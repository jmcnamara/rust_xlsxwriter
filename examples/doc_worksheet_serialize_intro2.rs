// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet. This version uses header deserialization.

use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
use serde::{Deserialize, Serialize};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some formats to use with the serialization data.
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_background_color("C6E0B4");

    // Create a serializable struct.
    #[derive(Deserialize, Serialize)]
    #[serde(rename_all = "PascalCase")]
    struct Student<'a> {
        name: &'a str,
        age: u8,
        id: u32,
    }

    let students = [
        Student {
            name: "Aoife",
            age: 25,
            id: 564351,
        },
        Student {
            name: "Caoimhe",
            age: 21,
            id: 443287,
        },
    ];

    // Set up the start location and headers of the data to be serialized.
    worksheet.deserialize_headers_with_format::<Student>(1, 3, &header_format)?;

    // Serialize the data.
    worksheet.serialize(&students)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
