// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of formatting the chart rotation for pie and doughnut charts.

use rust_xlsxwriter::{Chart, ChartFormat, ChartPoint, ChartSolidFill, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 15)?;
    worksheet.write(1, 0, 15)?;
    worksheet.write(2, 0, 30)?;

    // Some point object with formatting to use in the Pie chart.
    let points = vec![
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
        ),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFC000")),
        ),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
        ),
    ];

    // Create a new chart.
    let mut chart = Chart::new_pie();

    // Add a data series with formatting.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$3")
        .set_points(&points);

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
