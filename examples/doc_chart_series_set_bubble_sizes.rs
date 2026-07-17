// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of creating a Bubble chart with bubble sizes.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write_column(0, 0, [1, 2, 3])?;
    worksheet.write_column(0, 1, [10, 40, 30])?;
    worksheet.write_column(0, 2, [5, 12, 8])?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Bubble);

    // Add a data series with the X values, Y values and bubble sizes.
    chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$3")
        .set_values("Sheet1!$B$1:$B$3")
        .set_bubble_sizes("Sheet1!$C$1:$C$3");

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 4, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
