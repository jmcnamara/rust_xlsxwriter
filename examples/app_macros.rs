// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of adding macros to an `rust_xlsxwriter` file using a VBA macros
//! file extracted from an existing Excel xlsm file.
//!
//! The `vba_extract` utility (https://crates.io/crates/vba_extract) can be used
//! to extract the `vbaProject.bin` file.

use rust_xlsxwriter::{Button, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add the VBA macro file.
    workbook.add_vba_project("examples/vbaProject.bin")?;

    // Add a worksheet and some text.
    let worksheet = workbook.add_worksheet();

    // Widen the first column for clarity.
    worksheet.set_column_width(0, 30)?;

    worksheet.write(2, 0, "Press the button to say hello:")?;

    // Add a button tied to a macro in the VBA project.
    let button = Button::new()
        .set_caption("Press Me")
        .set_macro("say_hello")
        .set_width(80)
        .set_height(30);

    worksheet.insert_button(2, 1, &button)?;

    // Save the file to disk. Note the `.xlsm` extension. This is required by
    // Excel or it raise a warning.
    workbook.save("macros.xlsm")?;

    Ok(())
}
