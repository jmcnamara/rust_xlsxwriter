// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of adding data labels to a chart series with number formatting.
//!
use rust_xlsxwriter::{Chart, ChartDataLabel, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 0.1)?;
    worksheet.write(1, 0, 0.4)?;
    worksheet.write(2, 0, 0.5)?;
    worksheet.write(3, 0, 0.2)?;
    worksheet.write(4, 0, 0.1)?;
    worksheet.write(5, 0, 0.5)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Line);

    // Add a data series.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$6")
        .set_data_label(ChartDataLabel::new().show_value().set_num_format("0.00%"));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
