// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of setting the chart series gap and overlap. Note that it only
//! needs to be applied to one of the series in the chart.

use rust_xlsxwriter::*;

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add the worksheet data that the charts will refer to.
    let data = [[105, 150, 130, 90], [50, 120, 100, 110]];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32, col_num as u16, *row_data)?;
        }
    }

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure the data series and add a gap/overlap. Note that it only needs
    // to be applied to one of the series in the chart.
    chart
        .add_series()
        .set_values("Sheet1!$A$1:$A$4")
        .set_overlap(37)
        .set_gap(70);

    chart.add_series().set_values("Sheet1!$B$1:$B$4");

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 3, &chart)?;

    workbook.save("chart.xlsx")?;

    Ok(())
}
