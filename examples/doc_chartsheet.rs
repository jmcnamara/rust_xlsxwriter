// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! A simple chartsheet example. A chart is placed on it own dedicated
//! worksheet.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 10)?;
    worksheet.write(1, 0, 60)?;
    worksheet.write(2, 0, 30)?;
    worksheet.write(3, 0, 10)?;
    worksheet.write(4, 0, 50)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series using Excel formula syntax to describe the range.
    chart.add_series().set_values("Sheet1!$A$1:$A$5");

    // Create a new chartsheet.
    let chartsheet = workbook.add_chartsheet();

    // Add the chart to the chartsheet.
    chartsheet.insert_chart(0, 0, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
