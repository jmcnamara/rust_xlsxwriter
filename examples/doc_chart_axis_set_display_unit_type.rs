// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the units of the Value/Y-axis.

use rust_xlsxwriter::{Chart, ChartAxisDisplayUnitType, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 6_000_000)?;
    worksheet.write(1, 0, 17_000_000)?;
    worksheet.write(2, 0, 23_000_000)?;
    worksheet.write(3, 0, 4_000_000)?;
    worksheet.write(4, 0, 12_000_000)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series using Excel formula syntax to describe the range.
    chart.add_series().set_values("Sheet1!$A$1:$A$5");

    // Set the units for the axis.
    chart
        .y_axis()
        .set_display_unit_type(ChartAxisDisplayUnitType::Millions);

    // Hide legend for clarity.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
