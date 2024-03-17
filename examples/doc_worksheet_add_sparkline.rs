// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates adding a sparkline to a worksheet.

use rust_xlsxwriter::{Sparkline, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some sample data to plot.
    worksheet.write_row(0, 0, [-2, 2, 3, -1, 0])?;

    // Create a default line sparkline that plots the 1D data range.
    let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 0, 4));

    // Add it to the worksheet.
    worksheet.add_sparkline(0, 5, &sparkline)?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
