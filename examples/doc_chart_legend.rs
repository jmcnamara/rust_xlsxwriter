// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of getting the chart legend object and setting some of its
//! properties.

use rust_xlsxwriter::{Chart, ChartLegendPosition, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 50)?;
    worksheet.write(1, 0, 30)?;
    worksheet.write(2, 0, 40)?;
    worksheet.write(0, 1, 30)?;
    worksheet.write(1, 1, 35)?;
    worksheet.write(2, 1, 45)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add data series using Excel formula syntax to describe the range.
    chart.add_series().set_values("Sheet1!$A$1:$A$3");
    chart.add_series().set_values("Sheet1!$B$1:$B$3");

    // Turn on the chart legend and place it at the bottom of the chart.
    chart.legend().set_position(ChartLegendPosition::Bottom);

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 3, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
