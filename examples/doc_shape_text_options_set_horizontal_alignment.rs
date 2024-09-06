// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding a Textbox shape and setting some of the
//! text option properties. This highlights the difference between horizontal
//! and vertical centering.
//!
use rust_xlsxwriter::{
    Shape, ShapeText, ShapeTextHorizontalAlignment, ShapeTextVerticalAlignment, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some text for the textbox.
    let text = ["Some text", "on", "several lines"].join("\n");

    // Create a textbox shape and add some text with horizontal alignment.
    let textbox = Shape::textbox().set_text(&text).set_text_options(
        &ShapeText::new().set_horizontal_alignment(ShapeTextHorizontalAlignment::Center),
    );

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Create a textbox shape and add some text with vertical alignment.
    let textbox = Shape::textbox().set_text(&text).set_text_options(
        &ShapeText::new().set_vertical_alignment(ShapeTextVerticalAlignment::TopCentered),
    );

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 5, &textbox)?;

    // Save the file to disk.
    workbook.save("shape.xlsx")?;

    Ok(())
}
