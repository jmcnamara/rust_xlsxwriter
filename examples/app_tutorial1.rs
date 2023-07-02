// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A simple program to write some data to an Excel spreadsheet using
//! rust_xlsxwriter. Part 1 of a tutorial.

use rust_xlsxwriter::{Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![("Rent", 2000), ("Gas", 200), ("Food", 500), ("Gym", 100)];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Iterate over the data and write it out row by row.
    let mut row = 0;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write(row, 1, expense.1)?;
        row += 1;
    }

    // Write a total using a formula.
    worksheet.write(row, 0, "Total")?;
    worksheet.write(row, 1, Formula::new("=SUM(B1:B4)"))?;

    // Save the file to disk.
    workbook.save("tutorial1.xlsx")?;

    Ok(())
}
