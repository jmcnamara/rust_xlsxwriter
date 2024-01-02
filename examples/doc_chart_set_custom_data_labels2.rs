// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of adding custom data labels to a chart series.
//!
//! This example shows how to get the data from cells. In Excel this is a single
//! command called "Value from Cells" but in `rust_xlsxwriter` it needs to be
//! broken down into a cell reference for each data label.
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
    worksheet.write(0, 1, "Asia")?;
    worksheet.write(1, 1, "Africa")?;
    worksheet.write(2, 1, "Europe")?;
    worksheet.write(3, 1, "Americas")?;
    worksheet.write(4, 1, "Oceania")?;
    worksheet.write(5, 1, "Antarctic")?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Line);

    // Create some custom data labels.
    let data_labels = [
        ChartDataLabel::new().set_value("=Sheet1!$B$1").to_custom(),
        ChartDataLabel::new().set_value("=Sheet1!$B$2").to_custom(),
        ChartDataLabel::new().set_value("=Sheet1!$B$3").to_custom(),
        ChartDataLabel::new().set_value("=Sheet1!$B$4").to_custom(),
        ChartDataLabel::new().set_value("=Sheet1!$B$5").to_custom(),
        ChartDataLabel::new().set_value("=Sheet1!$B$6").to_custom(),
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
