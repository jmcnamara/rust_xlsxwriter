// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure, including chrono datetimes, to a worksheet.

use chrono::NaiveDate;
use rust_xlsxwriter::{CustomSerializeHeader, Format, FormatBorder, Workbook, XlsxError};
use serde::Serialize;

use rust_xlsxwriter::utility::serialize_chrono_naive_to_excel;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the date column for clarity.
    worksheet.set_column_width(1, 12)?;

    // Add some formats to use with the serialization data.
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_background_color("C6E0B4");

    let date_format = Format::new().set_num_format("yyyy/mm/dd");

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct Student<'a> {
        name: &'a str,

        // Note, we add a `rust_xlsxwriter` function to serialize the date.
        #[serde(serialize_with = "serialize_chrono_naive_to_excel")]
        dob: NaiveDate,

        id: u32,
    }

    let students = [
        Student {
            name: "Aoife",
            dob: NaiveDate::from_ymd_opt(1998, 1, 12).unwrap(),
            id: 564351,
        },
        Student {
            name: "Caoimhe",
            dob: NaiveDate::from_ymd_opt(2000, 5, 1).unwrap(),
            id: 443287,
        },
    ];

    // Set up the start location and headers of the data to be serialized. Note,
    // we need to add a cell format for the datetime data.
    let custom_headers = [
        CustomSerializeHeader::new("name")
            .rename("Student")
            .set_header_format(&header_format),
        CustomSerializeHeader::new("dob")
            .rename("Birthday")
            .set_cell_format(&date_format)
            .set_header_format(&header_format),
        CustomSerializeHeader::new("id")
            .rename("ID")
            .set_header_format(&header_format),
    ];

    worksheet.serialize_headers_with_options(0, 0, "Student", &custom_headers)?;

    // Serialize the data.
    worksheet.serialize(&students)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
