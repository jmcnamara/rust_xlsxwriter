// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{
    ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
};

// Create rust_xlsxwriter file to compare against Excel file.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");

    let conditional_format = ConditionalFormatCell::new()
        .set_rule(ConditionalFormatCellRule::EqualTo(false))
        .set_format(&format1);

    worksheet.add_conditional_format(8, 4, 8, 4, &conditional_format)?;

    worksheet.insert_checkbox(8, 4, false)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_checkbox05() {
    let test_runner = common::TestRunner::new()
        .set_name("checkbox05")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
