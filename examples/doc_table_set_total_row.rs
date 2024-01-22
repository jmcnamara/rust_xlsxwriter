// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of turning on the "totals" row at the bottom of a worksheet table.
//! Note, this just turns on the total run it doesn't add captions or subtotal
//! functions.

use rust_xlsxwriter::{Table, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some sample data for the table.
    let items = ["Apples", "Pears", "Bananas", "Oranges"];
    let data = [
        [10000, 5000, 8000, 6000],
        [2000, 3000, 4000, 5000],
        [6000, 6000, 6500, 6000],
        [500, 300, 200, 700],
    ];

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Set the column widths for clarity.
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Create a new table and configure the total row.
    let table = Table::new().set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 5, &table)?;

    // Save the file to disk.
    workbook.save("tables.xlsx")?;

    Ok(())
}
