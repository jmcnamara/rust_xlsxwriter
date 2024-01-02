// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

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

    // Set some variables to define the chart range.
    let row_min = 1;
    let row_max = values.len() as u32;
    let col_cat = 0;
    let col_val = 1;

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories(("Sheet1", row_min, col_cat, row_max, col_cat))
        .set_values(("Sheet1", row_min, col_val, row_max, col_val));

    // Add a chart title and some axis labels.
    chart.title().set_name("Results of sample tests");
    chart.x_axis().set_name("Test day");
    chart.y_axis().set_name("Sample length (mm)");

    // Turn off the chart legend.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(0, 2, &chart, 5, 5)?;

    workbook.save("chart_tutorial4.xlsx")?;

    Ok(())
}
