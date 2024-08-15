// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of creating a chart series as a standalone object and then adding
//! it to a chart via the [`Chart::push_series()`](Chart::add_series) method.

use rust_xlsxwriter::{Chart, ChartSeries, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 50)?;
    worksheet.write(1, 0, 30)?;
    worksheet.write(2, 0, 40)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Create a chart series and set the range for the values.
    let mut series = ChartSeries::new();
    series.set_values("Sheet1!$A$1:$A$3");

    // Add the data series to the chart.
    chart.push_series(&series);

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
