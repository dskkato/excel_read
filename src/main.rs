extern crate calamine;
use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};

fn example() -> Result<(), Error> {
    let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook
        .worksheet_range("Sheet1")
        .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;

    let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;

    if let Some(result) = iter.next() {
        let (label, value): (String, f64) = result?;
        assert_eq!(label, "celcius");
        assert_eq!(value, 22.2222);
        Ok(())
    } else {
        Err(From::from("expected at least one record but got none"))
    }
}

fn main() {
    println!("Hello, world!");
}
