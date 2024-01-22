// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of adding a header format to a column in a worksheet table.

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
    for col_num in 1..=6u16 {
        worksheet.set_column_width(col_num, 12)?;
    }

    // Create formats for the columns headers.
    let format1 = Format::new().set_font_color("#FF0000");
    let format2 = Format::new().set_font_color("#00FF00");
    let format3 = Format::new().set_font_color("#0000FF");
    let format4 = Format::new().set_font_color("#FFFF00");

    // Add a format to the columns headers.
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_header_format(format1),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_header_format(format2),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_header_format(format3),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_header_format(format4),
    ];

    // Create a new table and configure the columns.
    let table = Table::new().set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // Save the file to disk.
    workbook.save("tables.xlsx")?;

    Ok(())
}
