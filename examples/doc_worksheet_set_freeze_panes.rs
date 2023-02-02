// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the worksheet panes.

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let mut worksheet1 = Worksheet::new();
    let mut worksheet2 = Worksheet::new();
    let mut worksheet3 = Worksheet::new();

    worksheet1.write_string(0, 0, "Scroll down")?;
    worksheet2.write_string(0, 0, "Scroll across")?;
    worksheet3.write_string(0, 0, "Scroll down or across")?;

    // Freeze the top row only.
    worksheet1.set_freeze_panes(1, 0)?;

    // Freeze the leftmost column only.
    worksheet2.set_freeze_panes(0, 1)?;

    // Freeze the top row and leftmost column.
    worksheet3.set_freeze_panes(1, 1)?;

    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);
    workbook.push_worksheet(worksheet3);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
