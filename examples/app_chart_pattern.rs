// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A example of creating column charts using the rust_xlsxwriter library.

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Shingle", &bold)?;
    worksheet.write_with_format(0, 1, "Brick", &bold)?;

    let data = [[105, 150, 130, 90], [50, 120, 100, 110]];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the first data series and add fill patterns.
    chart
        .add_series()
        .set_name("Sheet1!$A$1")
        .set_values("Sheet1!$A$2:$A$5");

    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_values("Sheet1!$B$2:$B$5");

    // Add a chart title and some axis labels.
    chart.title().set_name("Cladding types");
    chart.x_axis().set_name("Region");
    chart.y_axis().set_name("Number of houses");

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 3, &chart)?;

    workbook.save("chart_pattern.xlsx")?;

    Ok(())
}
