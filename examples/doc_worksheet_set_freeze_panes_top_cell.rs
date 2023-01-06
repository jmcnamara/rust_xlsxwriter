// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the worksheet panes and also
//! setting the topmost visible cell in the scrolled area.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    worksheet.write_string_only(0, 0, "Scroll down")?;

    // Freeze the top row only.
    worksheet.set_freeze_panes(1, 0)?;

    // Pre-scroll to the row 20.
    worksheet.set_freeze_panes_top_cell(19, 0)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
