// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a format to a column in a worksheet table.

use rust_xlsxwriter::{Format, Table, TableColumn, Workbook, XlsxError};

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

    // Create a number format for number columns in the table.
    let format = Format::new().set_num_format("$#,##0.00");

    // Add a format to the number/currency columns.
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Q1").set_format(&format),
        TableColumn::new().set_header("Q2").set_format(&format),
        TableColumn::new().set_header("Q3").set_format(&format),
        TableColumn::new().set_header("Q4").set_format(&format),
    ];

    // Create a new table and configure the columns.
    let table = Table::new().set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // Save the file to disk.
    workbook.save("tables.xlsx")?;

    Ok(())
}
