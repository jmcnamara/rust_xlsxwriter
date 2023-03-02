// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A example of creating doughnut charts using the rust_xlsxwriter library.

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Category", &bold)?;
    worksheet.write_with_format(0, 1, "Values", &bold)?;

    worksheet.write(1, 0, "Glazed")?;
    worksheet.write(2, 0, "Chocolate")?;
    worksheet.write(3, 0, "Cream")?;

    worksheet.write(1, 1, 50)?;
    worksheet.write(2, 1, 35)?;
    worksheet.write(3, 1, 15)?;

    // -----------------------------------------------------------------------
    // Create a new doughnut chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Doughnut);

    // Configure the first data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4")
        .set_name("Doughnut sales data");

    // Add a chart title.
    chart.title().set_name("Popular Doughnut Types");

    // Set an Excel chart style.
    chart.set_style(10);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    workbook.save("chart_doughnut.xlsx")?;

    Ok(())
}
