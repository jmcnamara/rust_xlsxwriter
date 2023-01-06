// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A simple program to write some data to an Excel spreadsheet using
//! rust_xlsxwriter. Part 2 of a tutorial.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![("Rent", 2000), ("Gas", 200), ("Food", 500), ("Gym", 100)];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some column headers.
    worksheet.write_string(0, 0, "Item", &bold)?;
    worksheet.write_string(0, 1, "Cost", &bold)?;

    // Iterate over the data and write it out row by row.
    let mut row = 1;
    for expense in expenses.iter() {
        worksheet.write_string_only(row, 0, expense.0)?;
        worksheet.write_number(row, 1, expense.1, &money_format)?;
        row += 1;
    }

    // Write a total using a formula.
    worksheet.write_string(row, 0, "Total", &bold)?;
    worksheet.write_formula(row, 1, "=SUM(B2:B5)", &money_format)?;

    // Save the file to disk.
    workbook.save("tutorial2.xlsx")?;

    Ok(())
}
