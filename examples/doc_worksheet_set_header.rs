// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the worksheet header.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.set_header("&CPage &P of &N");

    worksheet.write_string_only(0, 0, "Hello")?;
    worksheet.write_string_only(200, 0, "Hello")?;
    worksheet.set_view_page_layout();

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
