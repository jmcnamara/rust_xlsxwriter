// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};

mod common;

// Test to demonstrate charts.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Add some test data for the chart(s).
    for row_num in 0..8 {
        for col_num in 0..6 {
            worksheet.write_number(row_num as u32, col_num as u16, 1)?;
        }
    }

    let mut chart = Chart::new(ChartType::LineStacked);
    chart.set_axis_ids(70523520, 77480704);

    chart.add_series().set_values(("Sheet1", 0, 0, 7, 0));
    chart.add_series().set_values(("Sheet1", 0, 1, 7, 1));
    chart.add_series().set_values(("Sheet1", 0, 2, 7, 2));
    chart.add_series().set_values(("Sheet1", 0, 3, 7, 3));
    chart.add_series().set_values(("Sheet1", 0, 4, 7, 4));
    chart.add_series().set_values(("Sheet1", 0, 5, 7, 5));

    worksheet.insert_chart(8, 7, &chart)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_chart_theme01() {
    let test_runner = common::TestRunner::new()
        .set_name("chart_theme01")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}