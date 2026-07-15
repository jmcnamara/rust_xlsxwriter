// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of creating 3D charts using the `rust_xlsxwriter` library.

use rust_xlsxwriter::{Chart, ChartType, Format, Workbook, XlsxError};

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

    // -----------------------------------------------------------------------
    // Create a 3D column chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Column3D);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("3D Column Chart");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a 3D stacked bar chart with a custom 3D view.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Bar3DStacked);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("3D Stacked Bar Chart");

    // Rotate the 3D view of the chart.
    chart.set_view_3d_x_rotation(30).set_view_3d_y_rotation(40);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a 3D area chart. The series are plotted along the depth axis.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Area3D);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("3D Area Chart");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a 3D pie chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Pie3D);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    // Add a chart title.
    chart.title().set_name("3D Pie Chart");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(49, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a 3D line chart. The series are plotted along the depth axis.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Line3D);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("3D Line Chart");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(65, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a 3D surface chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Surface3D);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("3D Surface Chart");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(81, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a contour chart: a surface chart viewed from above.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Contour);

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_name("Sheet1!$B$1");

    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_name("Sheet1!$C$1");

    // Add a chart title.
    chart.title().set_name("Contour Chart");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(97, 3, &chart, 25, 10)?;

    workbook.save("chart_3d.xlsx")?;

    Ok(())
}
