// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Workbook, XlsxError};
use serde::Serialize;

// Test case for Serde serialization. First test isn't serialized.
fn create_new_xlsx_file_1(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Not serialized.
    worksheet.write(0, 0, "col1")?;
    worksheet.write(1, 0, 123)?;
    worksheet.write(0, 1, "col2")?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: Option<u8>,
        col2: Option<u8>,
    }

    let data = MyStruct {
        col1: Some(123),
        col2: None,
    };

    worksheet.write_serialize_headers(0, 0, &["col1", "col2"])?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_serde03_1() {
    let test_runner = common::TestRunner::new()
        .set_name("serde03")
        .set_function(create_new_xlsx_file_1)
        .unique("1")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_serde03_2() {
    let test_runner = common::TestRunner::new()
        .set_name("serde03")
        .set_function(create_new_xlsx_file_2)
        .unique("2")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
