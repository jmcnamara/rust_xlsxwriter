// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the chart series categories and
//! values.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, "Jan")?;
    worksheet.write(1, 0, "Feb")?;
    worksheet.write(2, 0, "Mar")?;
    worksheet.write(0, 1, 50)?;
    worksheet.write(1, 1, 30)?;
    worksheet.write(2, 1, 40)?;
    worksheet.write(0, 2, 30)?;
    worksheet.write(1, 2, 40)?;
    worksheet.write(2, 2, 50)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series using Excel formula syntax to describe the range.
    chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$3")
        .set_values("Sheet1!$B$1:$B$3");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range. This method is better when you need to create the
    // ranges programmatically to match the data range in the worksheet.
    chart
        .add_series()
        .set_categories(("Sheet1", 0, 1, 2, 1))
        .set_values(("Sheet1", 0, 2, 2, 2));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 4, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
