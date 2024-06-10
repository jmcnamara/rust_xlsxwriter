// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of creating a combined Column and Line chart. In this example
//! they share the same primary Y axis.

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Add the worksheet data that the charts will refer to.
    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    worksheet.write_column_matrix(0, 0, data)?;

    // Create a new Column chart. This will be the primary chart.
    let mut column_chart = Chart::new(ChartType::Column);

    // Configure the data series for the primary chart.
    column_chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$6")
        .set_values("Sheet1!$B$1:$B$6");

    // Create a new line chart. This will use this as the secondary chart.
    let mut line_chart = Chart::new(ChartType::Line);

    // Configure the data series for the secondary chart.
    line_chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$6")
        .set_values("Sheet1!$C$1:$C$6");

    // Combine the charts.
    column_chart.combine(&line_chart);

    // Add some axis labels. Note, this is done via the primary chart.
    column_chart.x_axis().set_name("X axis");
    column_chart.y_axis().set_name("Y axis");

    // Add the primary chart to the worksheet.
    worksheet.insert_chart_with_offset(0, 3, &column_chart, 5, 5)?;

    // Save the file to disk.
    workbook.save("chart.xlsx")?;

    Ok(())
}
