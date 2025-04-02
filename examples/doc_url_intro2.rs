// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a URL to a worksheet.

use rust_xlsxwriter::{Url, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a URL with a Url struct.
    worksheet.write_url(0, 0, Url::new("https://www.rust-lang.org"))?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
