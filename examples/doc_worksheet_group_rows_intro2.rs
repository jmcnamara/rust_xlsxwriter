// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of how to group worksheet rows into outlines.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add grouping over the sub-total range.
    worksheet.group_rows(1, 10)?;

    // Add Level 1 grouping over the sub-total range.
    worksheet.group_rows(1, 4)?;
    worksheet.group_rows(6, 9)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
