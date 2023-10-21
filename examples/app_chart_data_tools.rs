// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! A demo of the various Excel chart data tools that are available via the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{
    Chart, ChartDataLabel, ChartErrorBars, ChartErrorBarsType, ChartMarker, ChartSolidFill,
    ChartTrendline, ChartTrendlineType, ChartType, Format, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let bold = Format::new().set_bold();

    // Add the worksheet data that the charts will refer to.
    worksheet.write_with_format(0, 0, "Number", &bold)?;
    worksheet.write_with_format(0, 1, "Data 1", &bold)?;
    worksheet.write_with_format(0, 2, "Data 2", &bold)?;

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
    // Trendline example
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series with a polynomial trendline.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_trendline(ChartTrendline::new().set_type(ChartTrendlineType::Polynomial(3)));

    // Configure the second series with a linear trendline.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7")
        .set_trendline(ChartTrendline::new().set_type(ChartTrendlineType::Linear));

    // Add a chart title.
    chart.title().set_name("Chart with Trendlines");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Data Labels and Markers example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series with data labels and markers.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_data_label(ChartDataLabel::new().show_value())
        .set_marker(ChartMarker::new().set_automatic());

    // Configure the second series as default.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title.
    chart.title().set_name("Chart with Data Labels and Markers");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Error Bar example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series with error bars.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7")
        .set_y_error_bars(ChartErrorBars::new().set_type(ChartErrorBarsType::StandardError));

    // Configure the second series as default.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title.
    chart.title().set_name("Chart with Error Bars");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(33, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Up-Down Bar example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add the chart up-down bars.
    chart
        .set_up_down_bars(true)
        .set_up_bar_format(ChartSolidFill::new().set_color("#00B050"))
        .set_down_bar_format(ChartSolidFill::new().set_color("#FF0000"));

    // Add a chart title.
    chart.title().set_name("Chart with Up-Down Bars");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(49, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // High-Low Lines example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add the chart High-Low lines.
    chart.set_high_low_lines(true);

    // Add a chart title.
    chart.title().set_name("Chart with High-Low Lines");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(65, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Drop Lines example.
    // -----------------------------------------------------------------------

    // Create a new Line chart.
    let mut chart = Chart::new(ChartType::Line);

    // Configure the first series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    // Configure the second series.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add the chart Drop lines.
    chart.set_drop_lines(true);

    // Add a chart title.
    chart.title().set_name("Chart with Drop Lines");

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(81, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("chart_data_tools.xlsx")?;

    Ok(())
}
