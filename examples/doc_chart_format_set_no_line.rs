// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of turning off a default line in a chart format.

use rust_xlsxwriter::{Chart, ChartFormat, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 1)?;
    worksheet.write(1, 0, 2)?;
    worksheet.write(2, 0, 3)?;
    worksheet.write(3, 0, 4)?;
    worksheet.write(4, 0, 5)?;
    worksheet.write(5, 0, 6)?;
    worksheet.write(0, 1, 10)?;
    worksheet.write(1, 1, 40)?;
    worksheet.write(2, 1, 50)?;
    worksheet.write(3, 1, 20)?;
    worksheet.write(4, 1, 10)?;
    worksheet.write(5, 1, 50)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::ScatterStraightWithMarkers);

    // Add a data series with formatting.
    chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$6")
        .set_values("Sheet1!$B$1:$B$6")
        .set_format(ChartFormat::new().set_no_line());

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
