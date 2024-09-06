// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example demonstrates adding a Textbox shape to a worksheet cell at an
//! offset.

use rust_xlsxwriter::{Shape, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a textbox shape and add some text.
    let textbox = Shape::textbox().set_text("This is some text");

    // Insert a textbox in a cell.
    worksheet.insert_shape_with_offset(1, 1, &textbox, 10, 5)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
