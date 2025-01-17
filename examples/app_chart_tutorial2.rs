// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of creating a simple chart using the rust_xlsxwriter library.

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    let categories = ["Mon", "Tue", "Wed", "Thu", "Fri"];
    let values = [20, 40, 50, 30, 20];

    worksheet.write_with_format(0, 0, "Day", &bold)?;
    worksheet.write_column(1, 0, categories)?;

    worksheet.write_with_format(0, 1, "Sample", &bold)?;
    worksheet.write_column(1, 1, values)?;

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$6")
        .set_values("Sheet1!$B$2:$B$6");

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    workbook.save("chart_tutorial2.xlsx")?;

    Ok(())
}
