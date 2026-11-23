Unicode True
RequestExecutionLevel user
SetCompressor /SOLID lzma

!ifndef STAGE
  !error "STAGE is required"
!endif
!ifndef OUTPUT
  !error "OUTPUT is required"
!endif
!ifndef ICON
  !error "ICON is required"
!endif

Name "LocalPDF"
Caption "LocalPDF Setup"
OutFile "${OUTPUT}"
InstallDir "$LOCALAPPDATA\LocalPDF"
InstallDirRegKey HKCU "Software\LocalPDF" "InstallLocation"
Icon "${ICON}"
UninstallIcon "${ICON}"
BrandingText "LocalPDF — private and offline"

VIProductVersion "1.0.0.0"
VIAddVersionKey "ProductName" "LocalPDF"
VIAddVersionKey "CompanyName" "LocalPDF"
VIAddVersionKey "FileDescription" "LocalPDF Setup"
VIAddVersionKey "FileVersion" "1.0.0"
VIAddVersionKey "ProductVersion" "1.0.0"
VIAddVersionKey "LegalCopyright" "LocalPDF"

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Function .onInit
  SetShellVarContext current
FunctionEnd

Function un.onInit
  SetShellVarContext current
FunctionEnd

Section "LocalPDF" MainSection
  SetOutPath "$INSTDIR"
  File /r "${STAGE}\*.*"

  CreateDirectory "$SMPROGRAMS\LocalPDF"
  CreateShortCut "$DESKTOP\LocalPDF.lnk" "$INSTDIR\LocalPDF.exe" "" "$INSTDIR\LocalPDF.exe" 0
  CreateShortCut "$SMPROGRAMS\LocalPDF\LocalPDF.lnk" "$INSTDIR\LocalPDF.exe" "" "$INSTDIR\LocalPDF.exe" 0
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  WriteRegStr HKCU "Software\LocalPDF" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "DisplayName" "LocalPDF"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "DisplayIcon" "$INSTDIR\LocalPDF.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "DisplayVersion" "1.0.0"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "Publisher" "LocalPDF"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF" "NoRepair" 1
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\LocalPDF.lnk"
  Delete "$SMPROGRAMS\LocalPDF\LocalPDF.lnk"
  RMDir "$SMPROGRAMS\LocalPDF"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\LocalPDF"
  DeleteRegKey HKCU "Software\LocalPDF"
  RMDir /r "$TEMP\LocalPDF"
  RMDir /r "$INSTDIR"
SectionEnd
