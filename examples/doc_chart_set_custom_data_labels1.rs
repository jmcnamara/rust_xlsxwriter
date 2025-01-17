// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of adding custom data labels to a chart series. This is useful
//! when you want to label the points of a data series with information that
//! isn't contained in the value or category names.
use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};

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

    // Create some custom data labels.
    let data_labels = [
        ChartDataLabel::new().set_value("Alice").to_custom(),
        ChartDataLabel::new().set_value("Bob").to_custom(),
        ChartDataLabel::new().set_value("Carol").to_custom(),
        ChartDataLabel::new().set_value("Dave").to_custom(),
        ChartDataLabel::new().set_value("Eve").to_custom(),
        ChartDataLabel::new().set_value("Frank").to_custom(),
    ];

    // Add a data series.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$6")
        .set_custom_data_labels(&data_labels);

    // Turn legend off for clarity.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
