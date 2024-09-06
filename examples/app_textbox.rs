// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Demonstrate adding a Textbox to a worksheet using the rust_xlsxwriter
//! library.

use rust_xlsxwriter::{Shape, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some text to add to the text box.
    let text = "This is an example of adding a textbox with some text in it";

    // Create a textbox shape and add the text.
    let textbox = Shape::textbox().set_text(text);

    // Insert a textbox in a cell.
    worksheet.insert_shape(1, 1, &textbox)?;

    // Save the file to disk.
    workbook.save("textbox.xlsx")?;

    Ok(())
}
