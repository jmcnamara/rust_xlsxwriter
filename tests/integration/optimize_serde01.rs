// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
use serde::{Deserialize, Serialize};

// Test case for Serde serialization. First test isn't serialized.
fn create_new_xlsx_file_1(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Not serialized.
    worksheet.write(0, 0, "col1")?;
    worksheet.write(0, 1, "col2")?;
    worksheet.write(1, 0, 1)?;
    worksheet.write(1, 1, -1)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. i8/u8.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: i8,
    }

    let data = MyStruct { col1: 1, col2: -1 };

    worksheet.serialize_headers(0, 0, &data)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. i16/u16.
fn create_new_xlsx_file_3(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u16,
        col2: i16,
    }

    let data = MyStruct { col1: 1, col2: -1 };

    worksheet.serialize_headers(0, 0, &data)?;

    worksheet.serialize(&data)?;
    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. i32/u32.
fn create_new_xlsx_file_4(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u32,
        col2: i32,
    }

    let data = MyStruct { col1: 1, col2: -1 };

    worksheet.serialize_headers(0, 0, &data)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. f32.
fn create_new_xlsx_file_5(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: f32,
        col2: f32,
    }

    let data = MyStruct {
        col1: 1.0,
        col2: -1.0,
    };

    worksheet.serialize_headers(0, 0, &data)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. f64.
fn create_new_xlsx_file_6(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: f64,
        col2: f32,
    }

    let data = MyStruct {
        col1: 1.0,
        col2: -1.0,
    };

    worksheet.serialize_headers(0, 0, &data)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. i64/u64.
fn create_new_xlsx_file_7(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u64,
        col2: i64,
    }

    let data = MyStruct { col1: 1, col2: -1 };

    worksheet.serialize_headers(0, 0, &data)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. Header deserialization.
fn create_new_xlsx_file_8(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Deserialize, Serialize)]
    struct MyStruct {
        col1: u8,
        col2: i8,
    }

    let data = MyStruct { col1: 1, col2: -1 };

    worksheet.deserialize_headers::<MyStruct>(0, 0)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. Proc Macro.
fn create_new_xlsx_file_9(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(XlsxSerialize, Serialize)]
    struct MyStruct {
        col1: u8,
        col2: i8,
    }

    let data = MyStruct { col1: 1, col2: -1 };

    worksheet.set_serialize_headers::<MyStruct>(0, 0)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_optimize_serde01_1() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_1)
        .unique("optimize1")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_2() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_2)
        .unique("optimize2")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_3() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_3)
        .unique("optimize3")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_4() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_4)
        .unique("optimize4")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_5() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_5)
        .unique("optimize5")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_6() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_6)
        .unique("optimize6")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_7() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_7)
        .unique("optimize7")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_8() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_8)
        .unique("optimize8")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde01_9() {
    let test_runner = common::TestRunner::new()
        .set_name("serde01")
        .set_function(create_new_xlsx_file_9)
        .unique("optimize9")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}