// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the zip compression level of a
//! workbook to a different level than the default, to make saving faster in
//! exchange for a slightly larger file.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Write some data to the worksheet.
    for row in 0..1000 {
        worksheet.write_number(row, 0, row)?;
        worksheet.write_string(row, 1, "Hello")?;
    }

    // Use a faster compression level than the default level 6.
    workbook.set_zip_compression_level(3);

    workbook.save("workbook.xlsx")?;

    Ok(())
}
