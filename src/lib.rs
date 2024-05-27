use office::{DataType,Range,Excel};
use csv::Writer;
use std::io::{self,Write};
use std::path::Path;

pub mod cli;

///Convert `office::DataType`s to `String`s
fn dt2string(dt: &DataType) -> String {
    match dt {
	DataType::Int(x) => x.to_string(),
	DataType::Float(x) => x.to_string(),
	DataType::String(x) => x.clone(),
	DataType::Bool(x) => if *x {String::from("true")} else {String::from("false")},
	DataType::Error(_) => String::from("NaN"),
	DataType::Empty => String::from(""),
    }
}

/// Write an `office::Range` of cells to a stream
fn write_range<T:Write>(mut w: Writer<T>, range: &Range){
    for r in range.rows() {
	w.write_record(r.into_iter()
		       .map(|x| dt2string(x)))
	    .expect("couldn't write sheet data");
    }
}

/// Convert a single file.
pub fn convert_file(filepath: &Path,
		    to_stdout: bool) {
    assert!(filepath.is_file(),"input is not a regular file");
    let mut e = Excel::open(filepath).expect("Couldn't open source file");
    let sheet_names = e.sheet_names().expect("Couldn't access sheets");
    let numsheets = sheet_names.len();
    for sn in sheet_names {
	let basename = filepath.file_stem().unwrap().
	    to_str().unwrap();
	let newname = if numsheets < 2 {
	    format!("{}.csv",basename)
	} else {
	    format!("{}_{}.csv",basename,sn)
	};
	if to_stdout {
	    // emit the sheet name if there's more than one
	    if numsheets < 2 {
		println!("{}_{}.csv",basename,sn);
	    }
	    let writer = Writer::from_writer(io::stdout());
	    write_range(writer,&e.worksheet_range(&sn)
			.expect("couldn't fetch sheet data"));
	} else {
	    let newpath = filepath.with_file_name(newname);
	    assert!(!newpath.exists(),"{} already exists",newpath.to_str().unwrap());
	    let writer = Writer::from_path(newpath)
		.expect("couldn't open file");
	    write_range(writer,&e.worksheet_range(&sn)
			    .expect("couldn't fetch sheet data"));	
	}
    }
}
