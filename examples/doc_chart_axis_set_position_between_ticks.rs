// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the axes data position relative to the
//! tick marks. Notice that by setting the data columns "on" the tick the first
//! and last columns are cut off by the plot area.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 5)?;
    worksheet.write(1, 0, 30)?;
    worksheet.write(2, 0, 40)?;
    worksheet.write(3, 0, 30)?;
    worksheet.write(4, 0, 5)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series using Excel formula syntax to describe the range.
    chart.add_series().set_values("Sheet1!$A$1:$A$5");

    // Set the axes data position relative to the tick marks.
    chart.x_axis().set_position_between_ticks(false);

    // Hide legend for clarity.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
