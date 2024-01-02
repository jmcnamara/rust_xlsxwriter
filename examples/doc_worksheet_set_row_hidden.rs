// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates hiding a worksheet row.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Hide row 2 (with zero indexing).
    worksheet.set_row_hidden(1)?;

    worksheet.write_string(2, 0, "Row 2 is hidden")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
