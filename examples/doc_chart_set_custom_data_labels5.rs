// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of adding custom data labels to a chart series.
//!
//! This example shows how to format some of the data labels and leave the rest
//! with the default formatting.
use rust_xlsxwriter::{
    Chart, ChartDataLabel, ChartFormat, ChartLine, ChartSolidFill, ChartType, Workbook, XlsxError,
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

    // Create some custom data labels.
    let data_labels = [
        ChartDataLabel::new()
            .set_value("Start")
            .set_format(
                ChartFormat::new()
                    .set_border(ChartLine::new().set_color("#FF0000"))
                    .set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
            )
            .to_custom(),
        ChartDataLabel::new().set_hidden().to_custom(),
        ChartDataLabel::new().set_hidden().to_custom(),
        ChartDataLabel::new().set_hidden().to_custom(),
        ChartDataLabel::new().set_hidden().to_custom(),
        ChartDataLabel::new()
            .set_value("End")
            .set_format(
                ChartFormat::new()
                    .set_border(ChartLine::new().set_color("#FF0000"))
                    .set_solid_fill(ChartSolidFill::new().set_color("#FFFF00")),
            )
            .to_custom(),
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
