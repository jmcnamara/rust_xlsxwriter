// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates a static function which generally returns
//! one value. Compare this with the dynamic function output of
//! `doc_working_with_formulas_dynamic_len.rs`.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write a a static function.
    worksheet.write_formula(0, 1, "=LEN(A1:A3)")?;

    // Write some data for the function to operate on.
    worksheet.write_string(0, 0, "Foo")?;
    worksheet.write_string(1, 0, "Food")?;
    worksheet.write_string(2, 0, "Frood")?;

    workbook.save("function_old.xlsx")?;

    Ok(())
}
