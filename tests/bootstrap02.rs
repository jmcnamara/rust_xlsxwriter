// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0 Copyright 2022, John McNamara,
// jmcnamara@cpan.org

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

mod common;

// Test case to demonstrate creating a basic file with 3 worksheets and no data.
fn create_new_xlsx_file1(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    _ = workbook.add_worksheet();
    _ = workbook.add_worksheet();
    _ = workbook.add_worksheet();

    workbook.save(filename)?;

    Ok(())
}

// Test case to demonstrate creating a basic file with 3 worksheets via
// Worksheet::new() and push_worksheet().
fn create_new_xlsx_file2(filename: &str) -> Result<(), XlsxError> {
    let mut worksheet1 = Worksheet::new();
    let mut worksheet2 = Worksheet::new();
    let mut worksheet3 = Worksheet::new();

    worksheet1.set_name("Sheet1")?;
    worksheet2.set_name("Sheet2")?;
    worksheet3.set_name("Sheet3")?;

    let mut workbook = Workbook::new();

    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);
    workbook.push_worksheet(worksheet3);

    workbook.save(filename)?;

    Ok(())
}

// Test case to demonstrate creating a basic file with 3 worksheets via
// Worksheet::new() and push_worksheet() + default sheet names.
fn create_new_xlsx_file3(filename: &str) -> Result<(), XlsxError> {
    let worksheet1 = Worksheet::new();
    let worksheet2 = Worksheet::new();
    let worksheet3 = Worksheet::new();

    let mut workbook = Workbook::new();

    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);
    workbook.push_worksheet(worksheet3);

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn bootstrap02_multiple_worksheets() {
    let test_runner = common::TestRunner::new("bootstrap02")
        .unique("1")
        .initialize();

    _ = create_new_xlsx_file1(test_runner.output_file());

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn bootstrap02_multiple_new_worksheets() {
    let test_runner = common::TestRunner::new("bootstrap02")
        .unique("2")
        .initialize();

    _ = create_new_xlsx_file2(test_runner.output_file());

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn bootstrap02_multiple_new_worksheets_default_names() {
    let test_runner = common::TestRunner::new("bootstrap02")
        .unique("3")
        .initialize();

    _ = create_new_xlsx_file3(test_runner.output_file());

    test_runner.assert_eq();
    test_runner.cleanup();
}
