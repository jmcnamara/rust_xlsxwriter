// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! A simple program to write some data to an Excel spreadsheet using
//! rust_xlsxwriter. Part 1 of a tutorial.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![("Rent", 2000), ("Gas", 200), ("Food", 500), ("Gym", 100)];

    // Create a new Excel file.
    let mut workbook = Workbook::new("tutorial1.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Iterate over the data and write it out row by row.
    let mut row = 0;
    for expense in expenses.iter() {
        worksheet.write_string_only(row, 0, expense.0)?;
        worksheet.write_number_only(row, 1, expense.1)?;
        row += 1;
    }

    // Write a total using a formula.
    worksheet.write_string_only(row, 0, "Total")?;
    worksheet.write_formula_only(row, 1, "=SUM(B1:B4)")?;

    // Close the file.
    workbook.close()?;

    Ok(())
}
