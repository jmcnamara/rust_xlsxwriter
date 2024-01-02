// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! A chart example demonstrating setting the maximum and minimum values for a
//! date axis.

use rust_xlsxwriter::{Chart, ChartType, ExcelDateTime, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // Adjust the date column width for clarity.
    worksheet.set_column_width(0, 11)?;

    // Add some data for the chart.
    let dates = [
        ExcelDateTime::parse_from_str("2024-01-01")?,
        ExcelDateTime::parse_from_str("2024-01-02")?,
        ExcelDateTime::parse_from_str("2024-01-03")?,
        ExcelDateTime::parse_from_str("2024-01-04")?,
        ExcelDateTime::parse_from_str("2024-01-05")?,
        ExcelDateTime::parse_from_str("2024-01-06")?,
        ExcelDateTime::parse_from_str("2024-01-07")?,
    ];
    let values = [27, 25, 19, 20, 18, 15, 19];

    worksheet.write_column_with_format(0, 0, dates, &date_format)?;
    worksheet.write_column(0, 1, values)?;

    // Create a new chart.
    let mut chart = Chart::new(ChartType::Column);

    chart
        .add_series()
        .set_categories(("Sheet1", 0, 0, 6, 0))
        .set_values(("Sheet1", 0, 1, 6, 1));

    // Set the axis as a date axis.
    chart.x_axis().set_date_axis(true);

    // Set the min and max date values for the chart.
    let min_date = ExcelDateTime::parse_from_str("2024-01-02")?;
    let max_date = ExcelDateTime::parse_from_str("2024-01-06")?;

    chart.x_axis().set_min_date(min_date);
    chart.x_axis().set_max_date(max_date);

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 3, &chart)?;

    // Save the file.
    workbook.save("chart.xlsx")?;

    Ok(())
}
