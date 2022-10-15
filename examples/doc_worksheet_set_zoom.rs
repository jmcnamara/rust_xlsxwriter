// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the worksheet zoom level.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new("worksheet.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.write_string_only(0, 0, "Hello")?;
    worksheet.set_zoom(200);

    workbook.close()?;

    Ok(())
}
