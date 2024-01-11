// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates serializing instances of a Serde derived
//! data structure, including datetimes, to a worksheet.

use rust_xlsxwriter::{
    CustomSerializeField, ExcelDateTime, Format, FormatBorder, SerializeFieldOptions, Workbook,
    XlsxError,
};
use serde::{Deserialize, Serialize};

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

    // Create a serializable struct.
    #[derive(Deserialize, Serialize)]
    struct Student<'a> {
        name: &'a str,
        dob: ExcelDateTime,
        id: u32,
    }

    let students = [
        Student {
            name: "Aoife",
            dob: ExcelDateTime::from_ymd(1998, 1, 12)?,
            id: 564351,
        },
        Student {
            name: "Caoimhe",
            dob: ExcelDateTime::from_ymd(2000, 5, 1)?,
            id: 443287,
        },
    ];

    // Set up the start location and headers of the data to be serialized. Note,
    // we need to add a cell format for the datetime data.
    let custom_headers = [
        CustomSerializeField::new("name").rename("Student"),
        CustomSerializeField::new("dob")
            .rename("Birthday")
            .set_value_format(date_format),
        CustomSerializeField::new("id").rename("ID"),
    ];
    let header_options = SerializeFieldOptions::new()
        .set_header_format(header_format)
        .set_custom_headers(&custom_headers);

    worksheet.deserialize_headers_with_options::<Student>(0, 0, &header_options)?;

    // Serialize the data.
    worksheet.serialize(&students)?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
