// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of creating a chart with gradient fills using the rust_xlsxwriter
//! library.

use rust_xlsxwriter::{
    Chart, ChartGradientFill, ChartGradientStop, ChartType, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Batch 1", &bold)?;
    worksheet.write_with_format(0, 2, "Batch 2", &bold)?;

    let data = [
        [2, 3, 4, 5, 6, 7],
        [10, 40, 50, 20, 10, 50],
        [30, 60, 70, 50, 40, 30],
    ];
    for (col_num, col_data) in data.iter().enumerate() {
        for (row_num, row_data) in col_data.iter().enumerate() {
            worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
        }
    }

    // Create a new column chart.
    let mut chart = Chart::new(ChartType::Column);

    //
    // Create a gradient profile to the first series.
    //
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1")
        .set_format(ChartGradientFill::new().set_gradient_stops(&[
            ChartGradientStop::new("#963735", 0),
            ChartGradientStop::new("#F1DCDB", 100),
        ]));

    //
    // Create a gradient profile to the second series.
    //
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1")
        .set_format(ChartGradientFill::new().set_gradient_stops(&[
            ChartGradientStop::new("#E36C0A", 0),
            ChartGradientStop::new("#FCEADA", 100),
        ]));

    //
    // Create a gradient profile and add it to chart plot area.
    //
    chart
        .plot_area()
        .set_format(ChartGradientFill::new().set_gradient_stops(&[
            ChartGradientStop::new("#FFEFD1", 0),
            ChartGradientStop::new("#F0EBD5", 50),
            ChartGradientStop::new("#B69F66", 100),
        ]));

    // Add some axis labels.
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Turn off the chart legend.
    chart.legend().set_hidden();

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(0, 3, &chart, 25, 10)?;

    workbook.save("chart_gradient.xlsx")?;

    Ok(())
}
