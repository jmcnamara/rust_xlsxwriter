// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates hiding a worksheet column.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Hide column B.
    worksheet.set_column_hidden(1)?;

    worksheet.write_string(0, 3, "Column B is hidden")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
