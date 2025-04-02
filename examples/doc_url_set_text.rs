// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a URL to a worksheet with
//! alternative text.

use rust_xlsxwriter::{Url, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a URL with a Url struct and alternative text.
    worksheet.write(
        0,
        0,
        Url::new("https://www.rust-lang.org").set_text("Learn Rust"),
    )?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
