// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of setting the percentage for the data labels of a chart series.
//! Usually this only applies to a Pie or Doughnut chart.

use rust_xlsxwriter::{Chart, ChartDataLabel, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 15)?;
    worksheet.write(1, 0, 15)?;
    worksheet.write(2, 0, 30)?;

    // Create a new chart.
    let mut chart = Chart::new_pie();

    // Add a data series.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$3")
        .set_data_label(ChartDataLabel::new().show_percentage());

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
