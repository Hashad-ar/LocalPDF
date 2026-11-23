use std::{
    collections::HashSet,
    ffi::c_void,
    fs,
    mem::size_of,
    path::{Path, PathBuf},
    ptr::{null, null_mut},
    time::{Duration, Instant},
};

use serde::Serialize;
use windows_sys::{
    Win32::{
        Foundation::{CloseHandle, HANDLE, SysAllocStringLen, SysFreeString, SysStringLen},
        System::{
            Com::{
                CLSCTX_LOCAL_SERVER, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoCreateInstance,
                CoInitializeEx, CoUninitialize, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
                DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            JobObjects::{
                AssignProcessToJobObject, CreateJobObjectW, JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
                JOBOBJECT_EXTENDED_LIMIT_INFORMATION, JobObjectExtendedLimitInformation,
                SetInformationJobObject,
            },
            Ole::DISPID_PROPERTYPUT,
            Threading::{OpenProcess, PROCESS_SET_QUOTA, PROCESS_TERMINATE},
            Variant::{VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I4},
        },
    },
    core::{GUID, HRESULT, PWSTR},
};

const IID_NULL: GUID = GUID {
    data1: 0,
    data2: 0,
    data3: 0,
    data4: [0; 8],
};
const IID_IDISPATCH: GUID = GUID {
    data1: 0x0002_0400,
    data2: 0,
    data3: 0,
    data4: [0xC0, 0, 0, 0, 0, 0, 0, 0x46],
};
const LOCALE_USER_DEFAULT: u32 = 0x0400;

#[repr(C)]
struct IDispatchVTable {
    query_interface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    get_type_info_count: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    get_type_info: unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HRESULT,
    get_ids_of_names: unsafe extern "system" fn(
        *mut c_void,
        *const GUID,
        *mut PWSTR,
        u32,
        u32,
        *mut i32,
    ) -> HRESULT,
    invoke: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *const GUID,
        u32,
        u16,
        *mut DISPPARAMS,
        *mut VARIANT,
        *mut EXCEPINFO,
        *mut u32,
    ) -> HRESULT,
}

struct Dispatch(*mut c_void);

impl Dispatch {
    fn create_word() -> Result<Self, HelperError> {
        let mut clsid = GUID::default();
        let prog_id = wide("Word.Application");
        let hr = unsafe { CLSIDFromProgID(prog_id.as_ptr(), &mut clsid) };
        check_hr(hr, "Microsoft Word is not registered on this PC.")?;

        let mut raw = null_mut();
        let hr = unsafe {
            CoCreateInstance(
                &clsid,
                null_mut(),
                CLSCTX_LOCAL_SERVER,
                &IID_IDISPATCH,
                &mut raw,
            )
        };
        check_hr(hr, "Microsoft Word could not be started.")?;
        if raw.is_null() {
            return Err(HelperError::new(
                "word_unavailable",
                "Microsoft Word did not return an automation object.",
            ));
        }
        Ok(Self(raw))
    }

    fn vtable(&self) -> &IDispatchVTable {
        unsafe { &***(&self.0 as *const *mut c_void as *const *const *const IDispatchVTable) }
    }

    fn member_id(&self, name: &str) -> Result<i32, HelperError> {
        let requested_name = name;
        let mut encoded_name = wide(requested_name);
        let mut name_pointer = encoded_name.as_mut_ptr();
        let mut member_id = 0_i32;
        let hr = unsafe {
            (self.vtable().get_ids_of_names)(
                self.0,
                &IID_NULL,
                &mut name_pointer,
                1,
                LOCALE_USER_DEFAULT,
                &mut member_id,
            )
        };
        check_hr(
            hr,
            &format!("Word does not expose the {requested_name} automation member."),
        )?;
        Ok(member_id)
    }

    fn invoke(
        &self,
        name: &str,
        flags: u16,
        mut arguments: Vec<OwnedVariant>,
    ) -> Result<OwnedVariant, HelperError> {
        let member_id = self.member_id(name)?;
        arguments.reverse();
        let mut named_argument = DISPID_PROPERTYPUT;
        let mut parameters = DISPPARAMS {
            rgvarg: if arguments.is_empty() {
                null_mut()
            } else {
                arguments.as_mut_ptr().cast::<VARIANT>()
            },
            rgdispidNamedArgs: if flags == DISPATCH_PROPERTYPUT {
                &mut named_argument
            } else {
                null_mut()
            },
            cArgs: arguments.len() as u32,
            cNamedArgs: u32::from(flags == DISPATCH_PROPERTYPUT),
        };
        let mut result = VARIANT::default();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = 0_u32;
        let hr = unsafe {
            (self.vtable().invoke)(
                self.0,
                member_id,
                &IID_NULL,
                LOCALE_USER_DEFAULT,
                flags,
                &mut parameters,
                &mut result,
                &mut exception,
                &mut argument_error,
            )
        };
        if hr < 0 {
            let description = take_exception_description(&mut exception);
            return Err(HelperError::new(
                "word_automation_failed",
                format!(
                    "Word automation member {name} failed (HRESULT 0x{:08X}, argument {argument_error}). {description}",
                    hr as u32
                ),
            ));
        }
        Ok(OwnedVariant(result))
    }

    fn get(&self, name: &str) -> Result<OwnedVariant, HelperError> {
        self.invoke(name, DISPATCH_PROPERTYGET, Vec::new())
    }

    fn set(&self, name: &str, value: OwnedVariant) -> Result<(), HelperError> {
        self.invoke(name, DISPATCH_PROPERTYPUT, vec![value])?;
        Ok(())
    }

    fn call(&self, name: &str, arguments: Vec<OwnedVariant>) -> Result<OwnedVariant, HelperError> {
        self.invoke(name, DISPATCH_METHOD, arguments)
    }
}

impl Drop for Dispatch {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe {
                (self.vtable().release)(self.0);
            }
            self.0 = null_mut();
        }
    }
}

#[repr(transparent)]
struct OwnedVariant(VARIANT);

impl OwnedVariant {
    fn integer(value: i32) -> Self {
        let mut variant = VARIANT::default();
        variant.Anonymous.Anonymous.vt = VT_I4;
        variant.Anonymous.Anonymous.Anonymous.lVal = value;
        Self(variant)
    }

    fn boolean(value: bool) -> Self {
        let mut variant = VARIANT::default();
        variant.Anonymous.Anonymous.vt = VT_BOOL;
        variant.Anonymous.Anonymous.Anonymous.boolVal = if value { -1 } else { 0 };
        Self(variant)
    }

    fn path(path: &Path) -> Result<Self, HelperError> {
        use std::os::windows::ffi::OsStrExt;
        let value: Vec<u16> = path.as_os_str().encode_wide().collect();
        let bstr = unsafe { SysAllocStringLen(value.as_ptr(), value.len() as u32) };
        if bstr.is_null() {
            return Err(HelperError::new(
                "out_of_memory",
                "Windows could not allocate a document path.",
            ));
        }
        let mut variant = VARIANT::default();
        variant.Anonymous.Anonymous.vt = VT_BSTR;
        variant.Anonymous.Anonymous.Anonymous.bstrVal = bstr;
        Ok(Self(variant))
    }

    fn into_dispatch(mut self) -> Result<Dispatch, HelperError> {
        unsafe {
            if self.0.Anonymous.Anonymous.vt != VT_DISPATCH {
                return Err(HelperError::new(
                    "word_automation_failed",
                    "Word returned an unexpected automation value.",
                ));
            }
            let pointer = self.0.Anonymous.Anonymous.Anonymous.pdispVal;
            self.0.Anonymous.Anonymous.vt = VT_EMPTY;
            if pointer.is_null() {
                return Err(HelperError::new(
                    "word_automation_failed",
                    "Word returned an empty automation object.",
                ));
            }
            Ok(Dispatch(pointer))
        }
    }
}

impl Drop for OwnedVariant {
    fn drop(&mut self) {
        unsafe {
            windows_sys::Win32::System::Variant::VariantClear(&mut self.0);
        }
    }
}

struct ComApartment;

impl ComApartment {
    fn initialize() -> Result<Self, HelperError> {
        let hr = unsafe { CoInitializeEx(null(), COINIT_APARTMENTTHREADED as u32) };
        check_hr(hr, "Windows COM could not be initialized.")?;
        Ok(Self)
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

struct WordKillJob(HANDLE);

impl WordKillJob {
    fn assign(word_pid: u32) -> Result<Self, HelperError> {
        unsafe {
            let handle = CreateJobObjectW(null(), null());
            if handle.is_null() {
                return Err(last_error("A cleanup job could not be created for Word."));
            }

            let mut limits = JOBOBJECT_EXTENDED_LIMIT_INFORMATION::default();
            limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
            if SetInformationJobObject(
                handle,
                JobObjectExtendedLimitInformation,
                (&limits as *const JOBOBJECT_EXTENDED_LIMIT_INFORMATION).cast(),
                size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
            ) == 0
            {
                CloseHandle(handle);
                return Err(last_error("Word cleanup limits could not be configured."));
            }

            let process = OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE, 0, word_pid);
            if process.is_null() {
                CloseHandle(handle);
                return Err(last_error(
                    "The Word automation process could not be tracked.",
                ));
            }
            let assigned = AssignProcessToJobObject(handle, process);
            CloseHandle(process);
            if assigned == 0 {
                CloseHandle(handle);
                return Err(last_error(
                    "The Word automation process could not be isolated.",
                ));
            }
            Ok(Self(handle))
        }
    }
}

impl Drop for WordKillJob {
    fn drop(&mut self) {
        unsafe { CloseHandle(self.0) };
    }
}

#[derive(Debug)]
struct HelperError {
    code: &'static str,
    detail: String,
}

impl HelperError {
    fn new(code: &'static str, detail: impl Into<String>) -> Self {
        Self {
            code,
            detail: detail.into(),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HelperResult<'a> {
    ok: bool,
    code: &'a str,
    detail: &'a str,
}

pub fn word_registered() -> bool {
    let mut clsid = GUID::default();
    let prog_id = wide("Word.Application");
    unsafe { CLSIDFromProgID(prog_id.as_ptr(), &mut clsid) >= 0 }
}

pub fn run_from_arguments(arguments: &[String]) -> i32 {
    if arguments.len() != 7 || arguments[1] != "--word-helper" {
        return 64;
    }
    let direction = arguments[2].as_str();
    let input = PathBuf::from(&arguments[3]);
    let output = PathBuf::from(&arguments[4]);
    let pid_file = PathBuf::from(&arguments[5]);
    let result_file = PathBuf::from(&arguments[6]);
    let preexisting = std::env::var("LOCALPDF_PREEXISTING_WORD_PIDS")
        .unwrap_or_default()
        .split(',')
        .filter_map(|value| value.parse::<u32>().ok())
        .collect::<HashSet<_>>();

    let result = convert(direction, &input, &output, &pid_file, &preexisting);
    let payload = match &result {
        Ok(()) => HelperResult {
            ok: true,
            code: "ok",
            detail: "",
        },
        Err(error) => HelperResult {
            ok: false,
            code: error.code,
            detail: &error.detail,
        },
    };
    let _ = fs::write(
        result_file,
        serde_json::to_vec(&payload).unwrap_or_else(|_| b"{\"ok\":false}".to_vec()),
    );
    i32::from(result.is_err())
}

fn convert(
    direction: &str,
    input: &Path,
    output: &Path,
    pid_file: &Path,
    preexisting: &HashSet<u32>,
) -> Result<(), HelperError> {
    let _apartment = ComApartment::initialize()?;
    let word = Dispatch::create_word()
        .map_err(|error| HelperError::new("word_unavailable", error.detail))?;
    let word_pid = detect_new_word_process(preexisting)?;
    if let Err(error) = fs::write(pid_file, word_pid.to_string()) {
        let _ = word.call("Quit", vec![OwnedVariant::integer(0)]);
        return Err(HelperError::new(
            "word_start_failed",
            format!("Could not record Word ownership: {error}"),
        ));
    }
    let cleanup_job = match WordKillJob::assign(word_pid) {
        Ok(job) => job,
        Err(error) => {
            let _ = word.call("Quit", vec![OwnedVariant::integer(0)]);
            return Err(error);
        }
    };

    let mut document: Option<Dispatch> = None;
    let operation = (|| {
        word.set("Visible", OwnedVariant::boolean(false))?;
        word.set("DisplayAlerts", OwnedVariant::integer(0))?;
        word.set("AutomationSecurity", OwnedVariant::integer(3))?;

        let documents = word.get("Documents")?.into_dispatch()?;
        let opened = documents
            .call(
                "Open",
                vec![
                    OwnedVariant::path(input)?,
                    OwnedVariant::boolean(false),
                    OwnedVariant::boolean(true),
                    OwnedVariant::boolean(false),
                ],
            )
            .map_err(|error| HelperError::new("word_open_failed", error.detail))?
            .into_dispatch()?;
        document = Some(opened);

        match direction {
            "word-to-pdf" => {
                document.as_ref().expect("document assigned").call(
                    "ExportAsFixedFormat",
                    vec![OwnedVariant::path(output)?, OwnedVariant::integer(17)],
                )?;
            }
            "pdf-to-word" => {
                document.as_ref().expect("document assigned").call(
                    "SaveAs2",
                    vec![OwnedVariant::path(output)?, OwnedVariant::integer(16)],
                )?;
            }
            _ => {
                return Err(HelperError::new(
                    "invalid_direction",
                    "Unknown conversion direction.",
                ));
            }
        }
        Ok(())
    })();

    if let Some(document) = document.take() {
        let _ = document.call("Close", vec![OwnedVariant::integer(0)]);
    }
    let _ = word.call("Quit", vec![OwnedVariant::integer(0)]);
    drop(word);
    drop(cleanup_job);

    operation?;
    if !output.is_file() {
        return Err(HelperError::new(
            "conversion_failed",
            "Word completed without creating the expected output file.",
        ));
    }
    Ok(())
}

fn detect_new_word_process(preexisting: &HashSet<u32>) -> Result<u32, HelperError> {
    let started = Instant::now();
    loop {
        let new_processes = crate::process::word_processes()
            .difference(preexisting)
            .copied()
            .collect::<Vec<_>>();
        if new_processes.len() == 1 {
            return Ok(new_processes[0]);
        }
        if new_processes.len() > 1 {
            return Err(HelperError::new(
                "word_start_failed",
                "More than one new Word process appeared, so LocalPDF could not determine safe ownership.",
            ));
        }
        if started.elapsed() >= Duration::from_secs(3) {
            return Err(HelperError::new(
                "word_busy",
                "Word did not create a separate automation process, so LocalPDF left existing Word processes untouched.",
            ));
        }
        std::thread::sleep(Duration::from_millis(50));
    }
}

fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(Some(0)).collect()
}

fn check_hr(hr: HRESULT, detail: &str) -> Result<(), HelperError> {
    if hr < 0 {
        Err(HelperError::new(
            "word_automation_failed",
            format!("{detail} HRESULT 0x{:08X}", hr as u32),
        ))
    } else {
        Ok(())
    }
}

fn last_error(detail: &str) -> HelperError {
    HelperError::new(
        "word_start_failed",
        format!("{detail} {}", std::io::Error::last_os_error()),
    )
}

fn take_exception_description(exception: &mut EXCEPINFO) -> String {
    unsafe {
        let description = bstr_to_string(exception.bstrDescription);
        if !exception.bstrSource.is_null() {
            SysFreeString(exception.bstrSource);
            exception.bstrSource = null();
        }
        if !exception.bstrDescription.is_null() {
            SysFreeString(exception.bstrDescription);
            exception.bstrDescription = null();
        }
        if !exception.bstrHelpFile.is_null() {
            SysFreeString(exception.bstrHelpFile);
            exception.bstrHelpFile = null();
        }
        description
    }
}

unsafe fn bstr_to_string(value: *const u16) -> String {
    if value.is_null() {
        return String::new();
    }
    let length = unsafe { SysStringLen(value) } as usize;
    String::from_utf16_lossy(unsafe { std::slice::from_raw_parts(value, length) })
}
