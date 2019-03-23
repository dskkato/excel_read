extern crate calamine;

use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .nth(1)
        .expect("Please provide an excel file to convert");
    let sheet = env::args()
        .nth(2)
        .expect("Expecting a sheet name as second argument");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") => (),
        _ => panic!("Expecting an .xlsx file"),
    }

    let dest = sce.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).expect("Failed to create output file"));
    let mut xl: Xlsx<_> = open_workbook(&sce).expect("Cannot find the file");
    let range = xl.worksheet_range(&sheet).unwrap().unwrap();

    write_range(&mut dest, &range).unwrap();
    write_range(&mut std::io::stdout(), &range).unwrap();
}

fn write_range(dest: &mut Write, range: &Range<DataType>) -> ::std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match c {
                DataType::Empty => Ok(()),
                DataType::String(s) => write!(dest, "{}", s),
                DataType::Float(f) => write!(dest, "{}", f),
                DataType::Int(i) => write!(dest, "{}", i),
                DataType::Error(e) => write!(dest, "{:?}", e),
                DataType::Bool(b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ",")?;
            }
        }
        writeln!(dest)?;
    }
    Ok(())
}
