#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

#[allow(dead_code)]
mod word_helper;
#[allow(dead_code)]
mod process;

fn main() {
    let arguments = std::env::args().collect::<Vec<_>>();
    std::process::exit(word_helper::run_from_arguments(&arguments));
}
