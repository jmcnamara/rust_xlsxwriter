// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Example of how to add tables to a worksheet using the rust_xlsxwriter
//! library.
//!
//! Tables in Excel are used to group rows and columns of data into a single
//! structure that can be referenced in a formula or formatted collectively.

use rust_xlsxwriter::{Table, TableColumn, TableFunction, TableStyle, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Some sample data for the table.
    let items = ["Apples", "Pears", "Bananas", "Oranges"];
    let data = [
        [10000, 5000, 8000, 6000],
        [2000, 3000, 4000, 5000],
        [6000, 6000, 6500, 6000],
        [500, 300, 200, 700],
    ];

    // -----------------------------------------------------------------------
    // Example 1. Default table with no data.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with no data.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Create a new table.
    let table = Table::new();

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 2. Default table with data.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with data.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create a new table.
    let table = Table::new();

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 3. Table without default autofilter.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table without default autofilter.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_autofilter(false);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 4. Table without default header row.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table without default header row.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_header_row(false);

    // Add the table to the worksheet.
    worksheet.add_table(3, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 5. Default table with "First Column" and "Last Column" options.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Default table with 'First Column' and 'Last Column' options.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_first_column(true).set_last_column(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 6. Table with banded columns but without default banded rows.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with banded columns but without default banded rows.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let table = Table::new().set_banded_rows(false).set_banded_columns(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 7. Table with user defined column headers.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with user defined column headers.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
    ];

    let table = Table::new().set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 5, &table)?;

    // -----------------------------------------------------------------------
    // Example 8. Table with user defined column headers, and formulas.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with user defined column headers, and formulas.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
        TableColumn::new()
            .set_header("Year")
            .set_formula("SUM(Table8[@[Quarter 1]:[Quarter 4]])"),
    ];

    let table = Table::new().set_columns(&columns);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 6, 6, &table)?;

    // -----------------------------------------------------------------------
    // Example 9. Table with totals row (but no caption or totals).
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with totals row (but no caption or totals).";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new().set_header("Product"),
        TableColumn::new().set_header("Quarter 1"),
        TableColumn::new().set_header("Quarter 2"),
        TableColumn::new().set_header("Quarter 3"),
        TableColumn::new().set_header("Quarter 4"),
        TableColumn::new()
            .set_header("Year")
            .set_formula("SUM(Table9[@[Quarter 1]:[Quarter 4]])"),
    ];

    let table = Table::new().set_columns(&columns).set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

    // -----------------------------------------------------------------------
    // Example 10. Table with totals row with user captions and functions.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with totals row with user captions and functions.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new()
            .set_header("Product")
            .set_total_label("Totals"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Year")
            .set_total_function(TableFunction::Sum)
            .set_formula("SUM(Table10[@[Quarter 1]:[Quarter 4]])"),
    ];

    let table = Table::new().set_columns(&columns).set_total_row(true);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

    // -----------------------------------------------------------------------
    // Example 11. Table with alternative Excel style.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with alternative Excel style.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new()
            .set_header("Product")
            .set_total_label("Totals"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Year")
            .set_total_function(TableFunction::Sum)
            .set_formula("SUM(Table11[@[Quarter 1]:[Quarter 4]])"),
    ];

    let table = Table::new()
        .set_columns(&columns)
        .set_total_row(true)
        .set_style(TableStyle::Light11);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

    // -----------------------------------------------------------------------
    // Example 12. Table with Excel style removed.
    // -----------------------------------------------------------------------

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let caption = "Table with Excel style removed.";

    // Set the column widths for clarity.
    worksheet.set_column_range_width(1, 6, 12)?;

    // Write the caption.
    worksheet.write(0, 1, caption)?;

    // Write the table data.
    worksheet.write_column(3, 1, items)?;
    worksheet.write_row_matrix(3, 2, data)?;

    // Create and configure a new table.
    let columns = vec![
        TableColumn::new()
            .set_header("Product")
            .set_total_label("Totals"),
        TableColumn::new()
            .set_header("Quarter 1")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 2")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 3")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Quarter 4")
            .set_total_function(TableFunction::Sum),
        TableColumn::new()
            .set_header("Year")
            .set_total_function(TableFunction::Sum)
            .set_formula("SUM(Table12[@[Quarter 1]:[Quarter 4]])"),
    ];

    let table = Table::new()
        .set_columns(&columns)
        .set_total_row(true)
        .set_style(TableStyle::None);

    // Add the table to the worksheet.
    worksheet.add_table(2, 1, 7, 6, &table)?;

    // Save the file to disk.
    workbook.save("tables.xlsx")?;

    Ok(())
}
