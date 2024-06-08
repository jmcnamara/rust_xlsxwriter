// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{
    Chart, ChartEmptyCells, ChartFormat, ChartPoint, ChartSolidFill, ChartType, Workbook, XlsxError,
};

// Create rust_xlsxwriter file to compare against Excel file.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Add some test data for the chart(s).
    worksheet.write(1, 7, "Donut")?;
    worksheet.write(1, 8, "Pie")?;
    worksheet.write_column(2, 7, [25, 50, 25, 100])?;
    worksheet.write_column(2, 8, [75, 1, 124])?;

    let mut chart_doughnut = Chart::new(ChartType::Doughnut);

    let points = vec![
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
        ),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FFC000")),
        ),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#00B050")),
        ),
        ChartPoint::new().set_format(ChartFormat::new().set_no_fill()),
    ];

    chart_doughnut
        .add_series()
        .set_values(("Sheet1", 2, 7, 5, 7))
        .set_name(("Sheet1", 1, 7))
        .set_points(&points);

    chart_doughnut.show_empty_cells_as(ChartEmptyCells::Gaps);
    chart_doughnut.legend().set_hidden();
    chart_doughnut.set_rotation(270);
    chart_doughnut.set_chart_area_format(ChartFormat::new().set_no_fill().set_no_border());

    let mut chart_pie = Chart::new(ChartType::Pie);

    let points = vec![
        ChartPoint::new().set_format(ChartFormat::new().set_no_fill()),
        ChartPoint::new().set_format(
            ChartFormat::new().set_solid_fill(ChartSolidFill::new().set_color("#FF0000")),
        ),
        ChartPoint::new().set_format(ChartFormat::new().set_no_fill()),
    ];

    chart_pie
        .add_series()
        .set_values(("Sheet1", 2, 8, 5, 8))
        .set_name(("Sheet1", 1, 8))
        .set_points(&points);

    chart_pie.set_rotation(270);

    chart_doughnut.combine(&chart_pie);
    worksheet.insert_chart(0, 0, &chart_doughnut)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_chart_combined11() {
    let test_runner = common::TestRunner::new()
        .set_name("chart_combined11")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
