// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates changing the default format for a
//! workbook.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Create a new default format for the workbook.
    let format = Format::new()
        .set_font_name("Aptos Narrow")
        .set_font_size(11);

    // Set the default format for the workbook.
    workbook.set_default_format(&format, 20, 64)?;

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some text to demonstrate the changed default format.
    worksheet.write(0, 0, "Hello")?;

    // Save the workbook to disk.
    workbook.save("workbook.xlsx")?;

    Ok(())
}
