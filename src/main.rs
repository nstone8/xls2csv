use xls2csv::cli;
use clap::Parser;
use std::path::Path;

fn main() {
    let arg = cli::Args::parse();
    xls2csv::convert_file(&Path::new(&arg.excelfile),
			  arg.stdout);
}
