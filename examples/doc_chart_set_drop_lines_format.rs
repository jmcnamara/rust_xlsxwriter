// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of setting drop lines for a chart, with formatting.

use rust_xlsxwriter::{Chart, ChartLine, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    let data = [[5, 10, 15], [4, 9, 13], [3, 8, 10], [2, 7, 6], [1, 6, 4]];

    for (row_num, row_data) in data.iter().enumerate() {
        for (col_num, col_data) in row_data.iter().enumerate() {
            worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
        }
    }

    // Create the chart.
    let mut chart = Chart::new(ChartType::Line);
    chart
        .add_series()
        .set_categories(("Sheet1", 0, 0, 4, 0))
        .set_values(("Sheet1", 0, 1, 4, 1));

    chart
        .add_series()
        .set_categories(("Sheet1", 0, 0, 4, 0))
        .set_values(("Sheet1", 0, 2, 4, 2));

    // Set the drop lines.
    chart
        .set_drop_lines(true)
        .set_drop_lines_format(ChartLine::new().set_color("#FF0000"));

    worksheet.insert_chart(0, 4, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
