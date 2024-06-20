// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of formatting the chart plot area of a chart. In Excel the plot
//! area is the area between the axes on which the chart series are plotted.

use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 10)?;
    worksheet.write(1, 0, 40)?;
    worksheet.write(2, 0, 50)?;
    worksheet.write(3, 0, 20)?;
    worksheet.write(4, 0, 10)?;
    worksheet.write(5, 0, 50)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series with formatting.
    chart.add_series().set_values("Sheet1!$A$1:$A$6");

    chart
        .plot_area()
        .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFFFB3")));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
