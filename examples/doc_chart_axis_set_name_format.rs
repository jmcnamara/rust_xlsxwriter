// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the formatting of the title of chart
//! axes.

use rust_xlsxwriter::{
    Chart, ChartFormat, ChartLine, ChartSolidFill, ChartType, Workbook, XlsxError,
};

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

    // Set the chart axis titles.
    chart
        .x_axis()
        .set_name("Formatted axis title")
        .set_name_format(
            ChartFormat::new()
                .set_border(ChartLine::new().set_color("#FF0000"))
                .set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
        );

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
