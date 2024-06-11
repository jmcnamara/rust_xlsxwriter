// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting a secondary Y axis.

use rust_xlsxwriter::{Chart, ChartLegendPosition, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_column(0, 0, [2, 3, 4, 5, 6, 7])?;
    worksheet.write_column(0, 1, [10, 40, 50, 20, 10, 50])?;

    // Create a new line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure a series that defaults to the primary axis.
    chart.add_series().set_values("Sheet1!$A$1:$A$6");

    // Configure another series with a secondary axis.
    chart
        .add_series()
        .set_values("Sheet1!$B$1:$B$6")
        .set_secondary_axis(true);

    // Add some axis labels.
    chart.y_axis().set_name("Y axis");
    chart.y2_axis().set_name("Y2 axis");

    // Move the legend to the bottom for clarity.
    chart.legend().set_position(ChartLegendPosition::Bottom);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(0, 2, &chart, 5, 5)?;

    workbook.save("chart.xlsx")?;

    Ok(())
}
