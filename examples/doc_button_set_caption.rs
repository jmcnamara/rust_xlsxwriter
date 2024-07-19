// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of adding an Excel Form Control button to a worksheet. This
//! example demonstrates setting the button caption.

use rust_xlsxwriter::{Button, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Add a button with a default caption.
    let button = Button::new();
    worksheet.insert_button(2, 1, &button)?;

    // Add a button with a user defined caption.
    let button = Button::new().set_caption("Press Me");
    worksheet.insert_button(4, 1, &button)?;

    // Save the file to disk.
    workbook.save("button.xlsx")?;

    Ok(())
}
