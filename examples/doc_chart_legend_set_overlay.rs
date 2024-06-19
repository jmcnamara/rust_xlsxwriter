// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of overlaying the chart legend on the plot area.

use rust_xlsxwriter::{Chart, ChartLegendPosition, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 50)?;
    worksheet.write(1, 0, 30)?;
    worksheet.write(2, 0, 40)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series using Excel formula syntax to describe the range.
    chart.add_series().set_values("Sheet1!$A$1:$A$3");

    // Turn on the chart legend and place it at the top of the chart.
    chart.legend().set_position(ChartLegendPosition::Top);

    // Overlay the chart legend on the plot area.
    chart.legend().set_overlay(true);

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
