// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting page breaks for a worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();
    worksheet.write_string_only(100, 100, "Test")?;

    // Set a page break at rows 20, 40 and 60.
    worksheet.set_page_breaks(&[20, 40, 60])?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
