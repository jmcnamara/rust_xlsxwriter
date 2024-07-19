// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of adding an Excel Form Control button to a worksheet. This
//! example demonstrates setting the button macro.

use rust_xlsxwriter::{Button, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add the VBA macro file.
    workbook.add_vba_project("examples/vbaProject.bin")?;

    // Add a worksheet and some text.
    let worksheet = workbook.add_worksheet();

    // Add a button tied to a macro in the VBA project.
    let button = Button::new().set_macro("say_hello");

    worksheet.insert_button(2, 1, &button)?;

    // Save the file to disk. Note the `.xlsm` extension.
    workbook.save("macros.xlsm")?;

    Ok(())
}
