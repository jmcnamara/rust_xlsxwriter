// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of getting the dimensions of some serialized data. In this example
//! we use the dimensions to set a conditional format range.

use rust_xlsxwriter::{
    ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a serializable struct.
    #[derive(Serialize)]
    #[serde(rename_all = "PascalCase")]
    struct MyStruct {
        col1: u8,
        col2: u8,
        col3: u8,
        col4: u8,
    }

    // Create some sample data.
    #[rustfmt::skip]
    let data = [
        MyStruct {col1: 34,  col2: 73, col3: 39, col4: 32},
        MyStruct {col1: 5,   col2: 24, col3: 1,  col4: 84},
        MyStruct {col1: 28,  col2: 79, col3: 97, col4: 13},
        MyStruct {col1: 27,  col2: 71, col3: 40, col4: 17},
        MyStruct {col1: 88,  col2: 25, col3: 33, col4: 23},
        MyStruct {col1: 23,  col2: 99, col3: 20, col4: 88},
        MyStruct {col1: 7,   col2: 57, col3: 88, col4: 28},
        MyStruct {col1: 53,  col2: 78, col3: 1,  col4: 96},
        MyStruct {col1: 60,  col2: 54, col3: 81, col4: 66},
        MyStruct {col1: 70,  col2: 5,  col3: 46, col4: 14},
    ];

    // Set the serialization location and headers.
    worksheet.serialize_headers(0, 0, &data[1])?;

    // Serialize the data.
    worksheet.serialize(&data)?;

    // Add a format. Green fill with dark green text.
    let format = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Create a conditional format.
    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::GreaterThanOrEqualTo(50))
        .set_format(format);

    // Get the range that the serialization applies to.
    let (min_row, min_col, max_row, max_col) = worksheet.get_serialize_dimensions("MyStruct")?;

    // Write the conditional format to the serialization area. Note, we add 1 to
    // the minimum row number to skip the headers.
    worksheet.add_conditional_format(
        min_row + 1,
        min_col,
        max_row,
        max_col,
        &conditional_format,
    )?;

    // Save the file.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
