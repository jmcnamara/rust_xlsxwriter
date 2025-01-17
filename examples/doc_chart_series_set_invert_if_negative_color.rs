// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the "Invert if negative" property and
//! associated color for a chart series. This also requires that you set a solid
//! fill color for the series.

use rust_xlsxwriter::{Chart, ChartSolidFill, ChartType, Workbook, XlsxError};

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

    // Add a data series and set the "Invert if negative" property and color.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$5")
        .set_format(ChartSolidFill::new().set_color("#4C9900"))
        .set_invert_if_negative_color("#FF6666");

    // Hide legend for clarity.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
