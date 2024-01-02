// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates a static function which generally returns
//! one value turned into a dynamic array function which returns a range of
//! values.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write a dynamic formula using a static function.
    worksheet.write_dynamic_array_formula(0, 1, 0, 1, "=LEN(A1:A3)")?;

    // Write some data for the function to operate on.
    worksheet.write_string(0, 0, "Foo")?;
    worksheet.write_string(1, 0, "Food")?;
    worksheet.write_string(2, 0, "Frood")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
