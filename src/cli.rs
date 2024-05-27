use clap::Parser;

/// Simple program to greet a person
#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
pub struct Args {
    /// file or directory to convert
    pub excelfile: String,

    /// if set, send output to standard output rather than a file
    #[arg(long)]
    pub stdout: bool
}
