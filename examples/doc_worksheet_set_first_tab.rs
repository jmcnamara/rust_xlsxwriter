// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the first visible worksheet tab
//! in a workbook with a large number of worksheets.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    for _ in 0..16 {
        workbook.add_worksheet();
    }

    // Make "Sheet13" the leftmost visible tab and "Sheet16" the active tab.
    workbook.worksheet_from_index(12)?.set_first_tab(true);
    workbook.worksheet_from_index(15)?.set_active(true);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
