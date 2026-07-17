// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of creating the Excel 2016+ "ChartEx" chart types: Waterfall,
//! Funnel, Histogram, Pareto, Box and Whisker, Treemap and Sunburst.

use rust_xlsxwriter::{
    Chart, ChartFormat, ChartParentLabelLayout, ChartSolidFill, ChartType, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // -----------------------------------------------------------------------
    // Waterfall chart.
    // -----------------------------------------------------------------------
    let categories = ["Start", "Q1", "Q2", "Q3", "Q4", "End"];
    let values = [100, 30, -20, 40, 25, 175];

    worksheet.write_column(0, 0, categories)?;
    worksheet.write_column(0, 1, values)?;

    let mut chart = Chart::new(ChartType::Waterfall);
    chart
        .add_series()
        .set_categories("Sheet1!$A$1:$A$6")
        .set_values("Sheet1!$B$1:$B$6");

    // Mark the "Start" and "End" columns as totals and hide the connector
    // lines.
    chart
        .set_waterfall_subtotals(&[0, 5])
        .set_waterfall_connector_lines(false)
        .title()
        .set_name("Waterfall");

    worksheet.insert_chart(0, 3, &chart)?;

    // -----------------------------------------------------------------------
    // Funnel chart.
    // -----------------------------------------------------------------------
    let categories = ["Leads", "Contacts", "Meetings", "Proposals", "Sales"];
    let values = [1000, 700, 400, 200, 100];

    worksheet.write_column(0, 8, categories)?;
    worksheet.write_column(0, 9, values)?;

    let mut chart = Chart::new(ChartType::Funnel);
    chart
        .add_series()
        .set_categories("Sheet1!$I$1:$I$5")
        .set_values("Sheet1!$J$1:$J$5");

    chart.title().set_name("Funnel");

    worksheet.insert_chart(15, 3, &chart)?;

    // -----------------------------------------------------------------------
    // Histogram chart.
    // -----------------------------------------------------------------------
    let values = [
        17, 22, 25, 31, 34, 42, 44, 48, 51, 55, 58, 61, 63, 67, 72, 78, 85, 91,
    ];

    worksheet.write_column(0, 12, values)?;

    let mut chart = Chart::new(ChartType::Histogram);
    chart.add_series().set_values("Sheet1!$M$1:$M$18");

    chart.set_histogram_bin_width(20.0);
    chart.title().set_name("Histogram");

    worksheet.insert_chart(30, 3, &chart)?;

    // -----------------------------------------------------------------------
    // Pareto chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Pareto);
    chart.add_series().set_values("Sheet1!$M$1:$M$18");

    chart.set_histogram_bin_count(5);
    chart.title().set_name("Pareto");

    worksheet.insert_chart(45, 3, &chart)?;

    // -----------------------------------------------------------------------
    // Box and Whisker chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::BoxWhisker);
    chart.add_series().set_values("Sheet1!$M$1:$M$18");

    chart.set_box_whisker_mean_line(true);
    chart.title().set_name("Box and Whisker");

    worksheet.insert_chart(60, 3, &chart)?;

    // -----------------------------------------------------------------------
    // Treemap chart.
    // -----------------------------------------------------------------------
    let parents = ["North", "", "", "South", "", ""];
    let children = ["Apple", "Pear", "Plum", "Apple", "Cherry", "Fig"];
    let values = [100, 50, 30, 80, 40, 60];

    worksheet.write_column(0, 15, parents)?;
    worksheet.write_column(0, 16, children)?;
    worksheet.write_column(0, 17, values)?;

    let mut chart = Chart::new(ChartType::Treemap);
    chart
        .add_series()
        .set_categories("Sheet1!$P$1:$Q$6")
        .set_values("Sheet1!$R$1:$R$6");

    chart.set_treemap_parent_labels(ChartParentLabelLayout::Banner);
    chart.title().set_name("Treemap");

    worksheet.insert_chart(75, 3, &chart)?;

    // -----------------------------------------------------------------------
    // Sunburst chart.
    // -----------------------------------------------------------------------
    let mut chart = Chart::new(ChartType::Sunburst);
    chart
        .add_series()
        .set_categories("Sheet1!$P$1:$Q$6")
        .set_values("Sheet1!$R$1:$R$6")
        .set_format(ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#4472C4")));

    chart.title().set_name("Sunburst");

    worksheet.insert_chart(90, 3, &chart)?;

    workbook.save("chartex_charts.xlsx")?;

    Ok(())
}
