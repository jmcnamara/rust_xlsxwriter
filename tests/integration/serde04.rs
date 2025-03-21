// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

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
    worksheet.write(2, 0, 456)?;
    worksheet.write(3, 0, 789)?;
    worksheet.write(0, 1, "col2")?;
    worksheet.write(1, 1, true)?;
    worksheet.write(2, 1, false)?;
    worksheet.write(3, 1, true)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. Array of data.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u16,
        col2: bool,
    }

    let data = [
        MyStruct {
            col1: 123,
            col2: true,
        },
        MyStruct {
            col1: 456,
            col2: false,
        },
        MyStruct {
            col1: 789,
            col2: true,
        },
    ];

    worksheet.serialize_headers(0, 0, &data[0])?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. Individual structs.
fn create_new_xlsx_file_3(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u16,
        col2: bool,
    }

    let data1 = MyStruct {
        col1: 123,
        col2: true,
    };

    let data2 = MyStruct {
        col1: 456,
        col2: false,
    };

    let data3 = MyStruct {
        col1: 789,
        col2: true,
    };

    worksheet.serialize_headers(0, 0, &data1)?;
    worksheet.serialize(&data1)?;
    worksheet.serialize(&data2)?;
    worksheet.serialize(&data3)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization with newtype variant enum values.
fn create_new_xlsx_file_4(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Define an enum for the test.
    #[derive(Serialize)]
    #[serde(rename_all = "lowercase")]
    enum MyEnum {
        Foo(u16),
        Bar(bool),
    }

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: MyEnum,
        col2: MyEnum,
    }

    let data = [
        MyStruct {
            col1: MyEnum::Foo(123),
            col2: MyEnum::Bar(true),
        },
        MyStruct {
            col1: MyEnum::Foo(456),
            col2: MyEnum::Bar(false),
        },
        MyStruct {
            col1: MyEnum::Foo(789),
            col2: MyEnum::Bar(true),
        },
    ];

    worksheet.serialize_headers(0, 0, &data[0])?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_serde04_1() {
    let test_runner = common::TestRunner::new()
        .set_name("serde04")
        .set_function(create_new_xlsx_file_1)
        .unique("1")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_serde04_2() {
    let test_runner = common::TestRunner::new()
        .set_name("serde04")
        .set_function(create_new_xlsx_file_2)
        .unique("2")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_serde04_3() {
    let test_runner = common::TestRunner::new()
        .set_name("serde04")
        .set_function(create_new_xlsx_file_3)
        .unique("3")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_serde04_4() {
    let test_runner = common::TestRunner::new()
        .set_name("serde04")
        .set_function(create_new_xlsx_file_4)
        .unique("4")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
