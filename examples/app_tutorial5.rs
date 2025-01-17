// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A simple program to write some data to an Excel spreadsheet using
//! rust_xlsxwriter. Part 5 of a tutorial.

use rust_xlsxwriter::{cell_range, Chart, ExcelDateTime, Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![
        ("Rent", 2000, "2022-09-01"),
        ("Gas", 200, "2022-09-05"),
        ("Food", 500, "2022-09-21"),
        ("Gym", 100, "2022-09-28"),
    ];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");

    // Add a number format for cells with dates.
    let date_format = Format::new().set_num_format("d mmm yyyy");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some column headers.
    worksheet.write_with_format(0, 0, "Item", &bold)?;
    worksheet.write_with_format(0, 1, "Cost", &bold)?;
    worksheet.write_with_format(0, 2, "Date", &bold)?;

    // Adjust the date column width for clarity.
    worksheet.set_column_width(2, 15)?;

    // Iterate over the data and write it out row by row.
    let mut row = 1;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write_with_format(row, 1, expense.1, &money_format)?;

        let date = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_with_format(row, 2, &date, &date_format)?;

        row += 1;
    }

    // For clarity, define some variables to use in the formula and chart
    // ranges. Row and column numbers are all zero-indexed.
    let first_row = 1; // Skip the header row.
    let last_row = first_row + (expenses.len() as u32) - 1;
    let item_col = 0;
    let cost_col = 1;

    // Write a total using a formula.
    worksheet.write_with_format(row, 0, "Total", &bold)?;

    let range = cell_range(first_row, cost_col, last_row, cost_col);
    let formula = format!("=SUM({range})");
    worksheet.write_with_format(row, 1, Formula::new(formula), &money_format)?;

    // Add a chart to display the expenses.
    let mut chart = Chart::new_pie();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories(("Sheet1", first_row, item_col, last_row, item_col))
        .set_values(("Sheet1", first_row, cost_col, last_row, cost_col));

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 4, &chart)?;

    // Save the file to disk.
    workbook.save("tutorial5.xlsx")?;

    Ok(())
}
