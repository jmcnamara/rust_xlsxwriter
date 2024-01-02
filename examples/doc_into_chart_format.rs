// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of passing chart formatting parameters via the
//! [`IntoChartFormat`] trait.

use rust_xlsxwriter::{Chart, ChartFormat, ChartSolidFill, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 10)?;
    worksheet.write(1, 0, 40)?;
    worksheet.write(2, 0, 50)?;
    worksheet.write(0, 1, 20)?;
    worksheet.write(1, 1, 10)?;
    worksheet.write(2, 1, 50)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add formatting via ChartFormat and a ChartSolidFill sub struct.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$3")
        .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#40EABB")));

    // Add formatting using a ChartSolidFill struct directly.
    chart
        .add_series()
        .set_values("Sheet1!$B$1:$B$3")
        .set_format(ChartSolidFill::new().set_color("#AAC3F2"));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
