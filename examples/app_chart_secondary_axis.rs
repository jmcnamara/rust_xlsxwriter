// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of creating an Excel Line chart with a secondary axis using the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Aliens", &bold)?;
    worksheet.write_with_format(0, 1, "Humans", &bold)?;
    worksheet.write_column(1, 0, [2, 3, 4, 5, 6, 7])?;
    worksheet.write_column(1, 1, [10, 40, 50, 20, 10, 50])?;

    // Create a new line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure a series with a secondary axis.
    chart
        .add_series()
        .set_name("Sheet1!$A$1")
        .set_values("Sheet1!$A$2:$A$7")
        .set_secondary_axis(true);

    // Configure another series that defaults to the primary axis.
    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_values("Sheet1!$B$2:$B$7");

    // Add a chart title and some axis labels.
    chart.title().set_name("Survey results");
    chart.x_axis().set_name("Days");
    chart.y_axis().set_name("Population");
    chart.y2_axis().set_name("Laser wounds");
    chart.y_axis().set_major_gridlines(false);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    workbook.save("chart_secondary_axis.xlsx")?;

    Ok(())
}
