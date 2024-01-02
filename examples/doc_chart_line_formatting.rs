// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of formatting a line/border in a chart element.

use rust_xlsxwriter::{
    Chart, ChartFormat, ChartLine, ChartLineDashType, ChartType, Workbook, XlsxError,
};

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
    let mut chart = Chart::new(ChartType::Line);

    // Add a data series with formatting.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$6")
        .set_format(
            ChartFormat::new().set_line(
                ChartLine::new()
                    .set_color("#FF9900")
                    .set_width(5.25)
                    .set_dash_type(ChartLineDashType::SquareDot)
                    .set_transparency(70),
            ),
        );

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
