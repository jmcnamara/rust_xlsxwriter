// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of turning on the last column highlighting property in a worksheet
//! table. This is normally off by default.

use rust_xlsxwriter::{Table, TableColumn, Workbook, XlsxError};

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
    worksheet.set_column_range_width(1, 6, 12)?;

    // Add a structured reference formula to the last column and set the header
    // caption. The last column in `add_table()` should be extended to account
    // for this extra column.
    let columns = vec![
        TableColumn::default(),
        TableColumn::default(),
        TableColumn::default(),
        TableColumn::default(),
        TableColumn::default(),
        TableColumn::new()
            .set_header("Totals")
            .set_formula("SUM(Table1[@[Column2]:[Column5]])"),
    ];

    // Create a new table and configure the last column highlighting.
    let table = Table::new().set_last_column(true).set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 6, &table)?;

    // Save the file to disk.
    workbook.save("tables.xlsx")?;

    Ok(())
}
