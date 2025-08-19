use anyhow::Error;
#[allow(unused_imports)]
use std::io;

#[macro_export]
macro_rules! unwrap_or_throw {
    ($result:expr) => {
        match $result {
            Err(error) => throw_windows_err(error),
            Ok(value) => value,
        }
    };
}

pub fn pause_windows() {
    #[cfg(windows)]
    {
        println!("\nPress Enter to exit...");
        let mut input = String::new();
        io::stdin().read_line(&mut input).unwrap();
    }
}

pub fn exit_gracefully(farewell: String) -> ! {
    println!("{}", farewell);
    pause_windows();
    std::process::exit(0);
}

pub fn throw_windows_err(err: Error) -> ! {
    eprintln!("{}", err);
    pause_windows();
    std::process::exit(1);
}
