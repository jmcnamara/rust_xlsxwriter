// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{
    CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError, XlsxSerialize,
};
use serde::Serialize;

// Test case for Serde serialization. First test isn't serialized.
fn create_new_xlsx_file_1(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Not serialized.
    worksheet.write(0, 0, "col1")?;
    worksheet.write(0, 1, "col2")?;
    worksheet.write(1, 0, 1)?;
    worksheet.write(1, 1, "aaa")?;
    worksheet.write(2, 0, 2)?;
    worksheet.write(2, 1, "bbb")?;
    worksheet.write(3, 0, 3)?;
    worksheet.write(3, 1, "ccc")?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
        },
    ];

    worksheet.serialize_headers(0, 0, &data.get(0).unwrap())?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for skipping fields.
fn create_new_xlsx_file_3(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    #[allow(dead_code)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        #[serde(skip_serializing)]
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    worksheet.serialize_headers(0, 0, &data.get(0).unwrap())?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for serialize_headers_with_options().
fn create_new_xlsx_file_4(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("col1"),
        CustomSerializeField::new("col2"),
    ];
    let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;

    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for skipping fields via serialize_headers_with_options().
fn create_new_xlsx_file_5(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("col1"),
        CustomSerializeField::new("col2"),
    ];
    let header_options = SerializeFieldOptions::new()
        .set_custom_headers(&custom_headers)
        .use_custom_headers_only(true);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for skipping fields via skip().
fn create_new_xlsx_file_6(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("col1"),
        CustomSerializeField::new("col2"),
        CustomSerializeField::new("col3").skip(true),
    ];
    let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for serialize_headers_with_options(). Field order is changed.
fn create_new_xlsx_file_7(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col2: &'static str,
        col1: u8,
    }

    let data = [
        MyStruct {
            col2: "aaa",
            col1: 1,
        },
        MyStruct {
            col2: "bbb",
            col1: 2,
        },
        MyStruct {
            col2: "ccc",
            col1: 3,
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("col1"),
        CustomSerializeField::new("col2"),
    ];
    let header_options = SerializeFieldOptions::new()
        .set_custom_headers(&custom_headers)
        .use_custom_headers_only(true);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for field rename via Serde.
fn create_new_xlsx_file_8(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        #[serde(rename = "col1")]
        field1: u8,

        #[serde(rename = "col2")]
        field2: &'static str,
    }

    let data = [
        MyStruct {
            field1: 1,
            field2: "aaa",
        },
        MyStruct {
            field1: 2,
            field2: "bbb",
        },
        MyStruct {
            field1: 3,
            field2: "ccc",
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("col1"),
        CustomSerializeField::new("col2"),
    ];
    let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for field rename via rename().
fn create_new_xlsx_file_9(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        field1: u8,
        field2: &'static str,
    }

    let data = [
        MyStruct {
            field1: 1,
            field2: "aaa",
        },
        MyStruct {
            field1: 2,
            field2: "bbb",
        },
        MyStruct {
            field1: 3,
            field2: "ccc",
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("field1").rename("col1"),
        CustomSerializeField::new("field2").rename("col2"),
    ];
    let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for skipping fields via skip() for single field only.
fn create_new_xlsx_file_10(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    let custom_headers = [CustomSerializeField::new("col3").skip(true)];
    let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for skipping fields via skip() (custom headers only)
fn create_new_xlsx_file_11(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    let custom_headers = [
        CustomSerializeField::new("col1"),
        CustomSerializeField::new("col2"),
        CustomSerializeField::new("col3").skip(true),
    ];
    let header_options = SerializeFieldOptions::new()
        .set_custom_headers(&custom_headers)
        .use_custom_headers_only(true);

    worksheet.serialize_headers_with_options(0, 0, &data[0], &header_options)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for skipping fields via proc macro.
fn create_new_xlsx_file_12(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize, XlsxSerialize)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        #[xlsx(skip)]
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    worksheet.set_serialize_headers::<MyStruct>(0, 0)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde skipping fields via proc macro.
fn create_new_xlsx_file_13(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize, XlsxSerialize)]
    #[allow(dead_code)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        #[serde(skip_serializing)]
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    worksheet.set_serialize_headers::<MyStruct>(0, 0)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde skipping fields via proc macro.
fn create_new_xlsx_file_14(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize, XlsxSerialize)]
    #[allow(dead_code)]
    struct MyStruct {
        col1: u8,
        col2: &'static str,
        #[serde(skip)]
        col3: bool,
    }

    let data = [
        MyStruct {
            col1: 1,
            col2: "aaa",
            col3: true,
        },
        MyStruct {
            col1: 2,
            col2: "bbb",
            col3: true,
        },
        MyStruct {
            col1: 3,
            col2: "ccc",
            col3: true,
        },
    ];

    worksheet.set_serialize_headers::<MyStruct>(0, 0)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for field rename via Serde and proc macro.
fn create_new_xlsx_file_15(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize, XlsxSerialize)]
    struct MyStruct {
        #[serde(rename = "col1")]
        field1: u8,

        #[serde(rename = "col2")]
        field2: &'static str,
    }

    let data = [
        MyStruct {
            field1: 1,
            field2: "aaa",
        },
        MyStruct {
            field1: 2,
            field2: "bbb",
        },
        MyStruct {
            field1: 3,
            field2: "ccc",
        },
    ];

    worksheet.set_serialize_headers::<MyStruct>(0, 0)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. Test Result types.
fn create_new_xlsx_file_16(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet_with_low_memory();

    // Create a serializable test struct.
    #[derive(Serialize)]
    struct MyStruct {
        col1: Result<u8, &'static str>,
        col2: Result<f64, &'static str>,
    }

    let data = [
        MyStruct {
            col1: Ok(1),
            col2: Err("aaa"),
        },
        MyStruct {
            col1: Ok(2),
            col2: Err("bbb"),
        },
        MyStruct {
            col1: Ok(3),
            col2: Err("ccc"),
        },
    ];

    worksheet.serialize_headers(0, 0, &data.get(0).unwrap())?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_optimize_serde07_1() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_1)
        .unique("optimize1")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_2() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_2)
        .unique("optimize2")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
#[test]
fn test_optimize_serde07_3() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_3)
        .unique("optimize3")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_4() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_4)
        .unique("optimize4")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_5() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_5)
        .unique("optimize5")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_6() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_6)
        .unique("optimize6")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_7() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_7)
        .unique("optimize7")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_8() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_8)
        .unique("optimize8")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_9() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_9)
        .unique("optimize9")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_10() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_10)
        .unique("optimize10")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_11() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_11)
        .unique("optimize11")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_12() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_12)
        .unique("optimize12")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_13() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_13)
        .unique("optimize13")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_14() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_14)
        .unique("optimize14")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_15() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_15)
        .unique("optimize15")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_optimize_serde07_16() {
    let test_runner = common::TestRunner::new()
        .set_name("serde07")
        .set_function(create_new_xlsx_file_16)
        .unique("optimize16")
        .ignore_worksheet_spans()
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
