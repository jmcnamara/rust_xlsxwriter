// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates "auto"-fitting the the width of a column
//! in Excel based on the maximum string width. See also the
//! [`Worksheet::autofit()`] command.
use rust_xlsxwriter::{autofit_cell_width, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some string data to write.
    let cities = ["Addis Ababa", "Buenos Aires", "Cairo", "Dhaka"];

    // Write the strings:
    worksheet.write_column(0, 0, cities)?;

    // Find the maximum column width in pixels.
    let max_width = cities.iter().map(|s| autofit_cell_width(s)).max().unwrap();

    // Set the column width as if it was auto-fitted.
    worksheet.set_column_auto_width(0, max_width)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
