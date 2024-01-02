// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the chart series name.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add some data for the chart.
    worksheet.write(0, 0, "Month")?;
    worksheet.write(1, 0, "Jan")?;
    worksheet.write(2, 0, "Feb")?;
    worksheet.write(3, 0, "Mar")?;
    worksheet.write(0, 1, "Total")?;
    worksheet.write(1, 1, 30)?;
    worksheet.write(2, 1, 20)?;
    worksheet.write(3, 1, 40)?;
    worksheet.write(0, 2, "Q1")?;
    worksheet.write(1, 2, 10)?;
    worksheet.write(2, 2, 10)?;
    worksheet.write(3, 2, 10)?;
    worksheet.write(0, 3, "Q2")?;
    worksheet.write(1, 3, 20)?;
    worksheet.write(2, 3, 10)?;
    worksheet.write(3, 3, 30)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    // Add a data series with a simple string name.
    chart
        .add_series()
        .set_name("Year to date")
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$B$2:$B$4");

    // Add a data series using Excel formula syntax to describe the range/name.
    chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$4")
        .set_values("Sheet1!$C$2:$C$4");

    // Add another data series to the chart using the alternative tuple syntax
    // to describe the range/name. This method is better when you need to create
    // the ranges programmatically to match the data range in the worksheet.
    chart
        .add_series()
        .set_name(("Sheet1", 0, 3))
        .set_categories(("Sheet1", 1, 0, 3, 0))
        .set_values(("Sheet1", 1, 3, 3, 3));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 5, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
