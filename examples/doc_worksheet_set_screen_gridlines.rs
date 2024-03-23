// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates turn off the worksheet worksheet screen
//! gridlines.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.write(0, 0, "Hello")?;

    // Turn off the screen gridlines.
    worksheet.set_screen_gridlines(false);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
