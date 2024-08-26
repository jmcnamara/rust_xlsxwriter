// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding a Textbox shape and setting some of the
//! text option properties.
//!
use rust_xlsxwriter::{Shape, ShapeText, ShapeTextDirection, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a textbox shape and add some text.
    let textbox = Shape::textbox()
        .set_text("古池や\n蛙飛び込む\n水の音")
        .set_text_options(&ShapeText::new().set_direction(ShapeTextDirection::Rotate90EastAsian));

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Save the file to disk.
    workbook.save("shape.xlsx")?;

    Ok(())
}
