// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{
    ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
};

// Create rust_xlsxwriter file to compare against Excel file.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_background_color("FF0000");
    let format2 = Format::new().set_background_color("92D050");

    worksheet.write(0, 0, 10)?;
    worksheet.write(1, 0, 20)?;
    worksheet.write(2, 0, 30)?;
    worksheet.write(3, 0, 40)?;

    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::LessThan(5))
        .set_format(format1)
        .set_stop_if_true(true);

    worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::GreaterThan(20))
        .set_format(format2)
        .set_stop_if_true(true);

    worksheet.add_conditional_format(0, 0, 0, 0, &conditional_format)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_cond_format15() {
    let test_runner = common::TestRunner::new()
        .set_name("cond_format15")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
