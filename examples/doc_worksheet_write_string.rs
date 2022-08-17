// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting different formatting for numbers
//! in an Excel worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("strings.xlsx");

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let italic_format = Format::new().set_italic();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some strings with formatting.
    worksheet.write_string(0, 0, "Hello", &bold_format)?;
    worksheet.write_string(1, 0, "שָׁלוֹם", &bold_format)?;
    worksheet.write_string(2, 0, "नमस्ते", &italic_format)?;
    worksheet.write_string(3, 0, "こんにちは", &italic_format)?;

    workbook.close()?;

    Ok(())
}
