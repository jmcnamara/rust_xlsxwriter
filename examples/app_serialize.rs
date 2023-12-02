// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! TODO.

use rust_xlsxwriter::{Workbook, XlsxError};
use serde::Serialize;

fn main() -> Result<(), XlsxError> {
    #[derive(Serialize)]
    struct Test {
        field_foo: bool,
        field_num: i8,
    }

    let test = Test {
        field_foo: true,
        field_num: 123,
    };

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.add_serialize_headers(1, 5, &["field_foo", "field_num"])?;

    worksheet.serialize(&test)?;
    worksheet.serialize(&test)?;

    // Save the file to disk.
    workbook.save("serialize.xlsx")?;

    Ok(())
}
