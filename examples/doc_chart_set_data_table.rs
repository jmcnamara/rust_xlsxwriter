// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of adding a data table to a chart.

use rust_xlsxwriter::{Chart, ChartDataTable, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    let data = [[1, 2, 3], [2, 4, 6], [3, 6, 9], [4, 8, 12], [5, 10, 15]];
    for (row_num, row_data) in data.iter().enumerate() {
        for (col_num, col_data) in row_data.iter().enumerate() {
            worksheet.write_number(row_num as u32, col_num as u16, *col_data)?;
        }
    }

    // Create a new chart.
    let mut chart = Chart::new_column();
    chart.add_series().set_values("Sheet1!$A$1:$A$5");
    chart.add_series().set_values("Sheet1!$B$1:$B$5");
    chart.add_series().set_values("Sheet1!$C$1:$C$5");

    // Add a default data table to the chart.
    let table = ChartDataTable::default();
    chart.set_data_table(&table);

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 4, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
