// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the "Invert if negative" property for
//! a chart series.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, 10)?;
    worksheet.write(1, 0, -5)?;
    worksheet.write(2, 0, 20)?;
    worksheet.write(3, 0, -15)?;
    worksheet.write(4, 0, 10)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series and set the "Invert if negative" property.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$5")
        .set_invert_if_negative();

    // Hide legend for clarity.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
