// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates efficiently hiding the unused rows in a
//! worksheet.
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some data.
    worksheet.write(0, 0, "First row")?;
    worksheet.write(6, 0, "Last row")?;

    // Set the row height for the blank rows so that they are "used".
    for row in 1..6 {
        worksheet.set_row_height(row, 15)?;
    }

    // Hide all the unused rows.
    worksheet.hide_unused_rows(true);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
