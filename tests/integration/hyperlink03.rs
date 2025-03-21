// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Format, Workbook, XlsxError};

// Test to demonstrate simple hyperlinks.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    let format = Format::default();

    worksheet1.write_url_with_options(0, 0, "http://www.perl.org/", "", "", Some(&format))?;
    worksheet1.write_url_with_options(3, 3, "http://www.perl.org/", "", "", Some(&format))?;
    worksheet1.write_url_with_options(7, 0, "http://www.perl.org/", "", "", Some(&format))?;
    worksheet1.write_url_with_options(5, 1, "http://www.cpan.org/", "", "", Some(&format))?;
    worksheet1.write_url_with_options(11, 5, "http://www.cpan.org/", "", "", Some(&format))?;

    let worksheet2 = workbook.add_worksheet();

    worksheet2.write_url_with_options(1, 2, "http://www.google.com/", "", "", Some(&format))?;
    worksheet2.write_url_with_options(4, 2, "http://www.cpan.org/", "", "", Some(&format))?;
    worksheet2.write_url_with_options(6, 2, "http://www.perl.org/", "", "", Some(&format))?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_hyperlink03() {
    let test_runner = common::TestRunner::new()
        .set_name("hyperlink03")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
