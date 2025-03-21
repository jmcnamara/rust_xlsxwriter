// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{
    CustomSerializeField, SerializeFieldOptions, Table, TableColumn, TableFunction, Workbook,
    XlsxError,
};
use serde::Serialize;

// Test case for Serde serialization. First test isn't serialized.
fn create_new_xlsx_file_1(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width(1, 10.288)?;
    worksheet.set_column_width(2, 10.288)?;
    worksheet.set_column_width(3, 10.288)?;
    worksheet.set_column_width(4, 10.288)?;
    worksheet.set_column_width(5, 10.288)?;
    worksheet.set_column_width(6, 10.288)?;
    worksheet.set_column_width(7, 10.288)?;
    worksheet.set_column_width(8, 10.288)?;
    worksheet.set_column_width(9, 10.288)?;
    worksheet.set_column_width(10, 10.288)?;

    worksheet.write(0, 0, "Column1")?;
    worksheet.write(0, 1, "Column2")?;
    worksheet.write(0, 2, "Column3")?;
    worksheet.write(0, 3, "Column4")?;
    worksheet.write(0, 4, "Column5")?;
    worksheet.write(0, 5, "Column6")?;
    worksheet.write(0, 6, "Column7")?;
    worksheet.write(0, 7, "Column8")?;
    worksheet.write(0, 8, "Column9")?;
    worksheet.write(0, 9, "Column10")?;
    worksheet.write(0, 10, "Total")?;

    worksheet.write(3, 1, 0)?;
    worksheet.write(3, 2, 0)?;
    worksheet.write(3, 3, 0)?;
    worksheet.write(3, 6, 0)?;
    worksheet.write(3, 7, 0)?;
    worksheet.write(3, 8, 0)?;
    worksheet.write(3, 9, 0)?;
    worksheet.write(3, 10, 0)?;
    worksheet.write(4, 1, 0)?;
    worksheet.write(4, 2, 0)?;
    worksheet.write(4, 3, 0)?;
    worksheet.write(4, 6, 0)?;
    worksheet.write(4, 7, 0)?;
    worksheet.write(4, 8, 0)?;
    worksheet.write(4, 9, 0)?;
    worksheet.write(4, 10, 0)?;

    let columns = vec![
        TableColumn::new().set_total_label("Total"),
        TableColumn::default(),
        TableColumn::new().set_total_function(TableFunction::Average),
        TableColumn::new().set_total_function(TableFunction::Count),
        TableColumn::new().set_total_function(TableFunction::CountNumbers),
        TableColumn::new().set_total_function(TableFunction::Max),
        TableColumn::new().set_total_function(TableFunction::Min),
        TableColumn::new().set_total_function(TableFunction::Sum),
        TableColumn::new().set_total_function(TableFunction::StdDev),
        TableColumn::new().set_total_function(TableFunction::Var),
    ];

    let table = Table::new().set_columns(&columns).set_total_row(true);

    worksheet.add_table(2, 1, 5, 10, &table)?;

    workbook.save(filename)?;

    Ok(())
}

// Test case for Serde serialization. Test Worksheet table.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write(0, 0, "Column1")?;
    worksheet.write(0, 1, "Column2")?;
    worksheet.write(0, 2, "Column3")?;
    worksheet.write(0, 3, "Column4")?;
    worksheet.write(0, 4, "Column5")?;
    worksheet.write(0, 5, "Column6")?;
    worksheet.write(0, 6, "Column7")?;
    worksheet.write(0, 7, "Column8")?;
    worksheet.write(0, 8, "Column9")?;
    worksheet.write(0, 9, "Column10")?;
    worksheet.write(0, 10, "Total")?;

    #[derive(Serialize)]
    #[serde(rename_all = "PascalCase")]
    struct MyStruct {
        column1: u8,
        column2: u8,
        column3: u8,
        column4: Option<u8>,
        column5: Option<u8>,
        column6: u8,
        column7: u8,
        column8: u8,
        column9: u8,
        column10: u8,
    }

    let data = MyStruct {
        column1: 0,
        column2: 0,
        column3: 0,
        column4: None,
        column5: None,
        column6: 0,
        column7: 0,
        column8: 0,
        column9: 0,
        column10: 0,
    };

    // Create a user defined table.
    let columns = vec![
        TableColumn::new().set_total_label("Total"),
        TableColumn::default(),
        TableColumn::new().set_total_function(TableFunction::Average),
        TableColumn::new().set_total_function(TableFunction::Count),
        TableColumn::new().set_total_function(TableFunction::CountNumbers),
        TableColumn::new().set_total_function(TableFunction::Max),
        TableColumn::new().set_total_function(TableFunction::Min),
        TableColumn::new().set_total_function(TableFunction::Sum),
        TableColumn::new().set_total_function(TableFunction::StdDev),
        TableColumn::new().set_total_function(TableFunction::Var),
    ];

    let table = Table::new().set_columns(&columns).set_total_row(true);

    let header_options = SerializeFieldOptions::new()
        .set_table(&table)
        .set_custom_headers(&[
            CustomSerializeField::new("Column1").set_column_width(10.288),
            CustomSerializeField::new("Column2").set_column_width(10.288),
            CustomSerializeField::new("Column3").set_column_width(10.288),
            CustomSerializeField::new("Column4").set_column_width(10.288),
            CustomSerializeField::new("Column5").set_column_width(10.288),
            CustomSerializeField::new("Column6").set_column_width(10.288),
            CustomSerializeField::new("Column7").set_column_width(10.288),
            CustomSerializeField::new("Column8").set_column_width(10.288),
            CustomSerializeField::new("Column9").set_column_width(10.288),
            CustomSerializeField::new("Column10").set_column_width(10.288),
        ]);

    worksheet.serialize_headers_with_options(2, 1, &data, &header_options)?;

    worksheet.serialize(&data)?;
    worksheet.serialize(&data)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_serde22_1() {
    let test_runner = common::TestRunner::new()
        .set_name("table09")
        .set_function(create_new_xlsx_file_1)
        .ignore_calc_chain()
        .unique("serde22_1")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn test_serde22_2() {
    let test_runner = common::TestRunner::new()
        .set_name("table09")
        .set_function(create_new_xlsx_file_2)
        .ignore_calc_chain()
        .unique("serde22_2")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
