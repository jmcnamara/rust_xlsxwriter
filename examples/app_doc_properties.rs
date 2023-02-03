// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of setting workbook document properties for a file created using
//! the rust_xlsxwriter library.

use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let properties = DocProperties::new()
        .set_title("This is an example spreadsheet")
        .set_subject("That demonstrates document properties")
        .set_author("A. Rust User")
        .set_manager("J. Alfred Prufrock")
        .set_company("Rust Solutions Inc")
        .set_category("Sample spreadsheets")
        .set_keywords("Sample, Example, Properties")
        .set_comment("Created with Rust and rust_xlsxwriter");

    workbook.set_properties(&properties);

    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(0, 30)?;
    worksheet.write_string(0, 0, "See File -> Info -> Properties")?;

    workbook.save("doc_properties.xlsx")?;

    Ok(())
}
