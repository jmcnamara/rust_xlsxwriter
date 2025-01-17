// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of creating Excel Column charts with data tables using the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{Chart, ChartDataTable, ChartType, Format, Workbook, XlsxError};

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
    // Create a column chart with a data table.
    // -----------------------------------------------------------------------

    // Create a new Column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure some data series.
    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    chart.title().set_name("Chart with Data Table");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set a default data table on the X-Axis.
    let table = ChartDataTable::default();
    chart.set_data_table(&table);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(1, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Create a column chart with a data table and legend keys.
    // -----------------------------------------------------------------------

    // Create a new Column chart.
    let mut chart = Chart::new(ChartType::Column);

    // Configure some data series.
    chart
        .add_series()
        .set_name("Sheet1!$B$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$B$2:$B$7");

    chart
        .add_series()
        .set_name("Sheet1!$C$1")
        .set_categories("Sheet1!$A$2:$A$7")
        .set_values("Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    chart.title().set_name("Data Table with legend keys");
    chart.x_axis().set_name("Test number");
    chart.y_axis().set_name("Sample length (mm)");

    // Set a data table on the X-Axis with the legend keys shown.
    let table = ChartDataTable::new().show_legend_keys(true);
    chart.set_data_table(&table);

    // Add the chart to the worksheet.
    worksheet.insert_chart_with_offset(17, 3, &chart, 25, 10)?;

    // -----------------------------------------------------------------------
    // Save and close the file.
    // -----------------------------------------------------------------------
    workbook.save("chart_data_table.xlsx")?;

    Ok(())
}
