// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the top and leftmost visible cell
//! in the worksheet. Often used in conjunction with `set_selection()` to
//! activate the same cell.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Set top-left cell to AA32.
    worksheet.set_top_left_cell(31, 26)?;

    // Also make this the active/selected cell.
    worksheet.set_selection(31, 26, 31, 26)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
