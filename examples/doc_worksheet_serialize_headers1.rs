// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct.
    #[derive(Serialize)]
    #[serde(rename_all = "PascalCase")]
    struct Student<'a> {
        name: &'a str,
        age: u8,
        id: u32,
    }

    let student = Student {
        name: "Aoife",
        age: 25,
        id: 564351,
    };

    // Set up the start location and headers of the data to be serialized using
    // any temporary or valid instance.
    worksheet.serialize_headers(2, 4, &student)?;

    // Serialize the data.
    worksheet.serialize(&student)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
