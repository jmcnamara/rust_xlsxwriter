// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the point where the axes will cross.

use rust_xlsxwriter::{Chart, ChartAxisCrossing, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, "North")?;
    worksheet.write(1, 0, "South")?;
    worksheet.write(2, 0, "East")?;
    worksheet.write(3, 0, "West")?;
    worksheet.write(0, 1, 10)?;
    worksheet.write(1, 1, 35)?;
    worksheet.write(2, 1, 40)?;
    worksheet.write(3, 1, 25)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series using Excel formula syntax to describe the range.
    chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$5")
        .set_values("Sheet1!$B$1:$B$5");

    // Set the X-axis crossing at a category index.
    chart
        .x_axis()
        .set_crossing(ChartAxisCrossing::CategoryNumber(3));

    // Set the Y-axis crossing at a value.
    chart
        .y_axis()
        .set_crossing(ChartAxisCrossing::AxisValue(20.0));

    // Hide legend for clarity.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
