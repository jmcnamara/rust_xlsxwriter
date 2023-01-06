// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a default format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Create a new default format.
    let format = Format::default();

    // These methods calls are equivalent.
    worksheet.write_string_only(0, 0, "Hello")?;
    worksheet.write_string(1, 0, "Hello", &format)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
