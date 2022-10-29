// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing an Excel "Future Function" with
//! an explicit prefix.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write a future function with explicit prefix.
    worksheet.write_formula_only(0, 0, "=_xlfn.STDEV.S(B1:B5)")?;

    // Write some data for the function to operate on.
    worksheet.write_number_only(0, 1, 1.23)?;
    worksheet.write_number_only(1, 1, 1.03)?;
    worksheet.write_number_only(2, 1, 1.20)?;
    worksheet.write_number_only(3, 1, 1.15)?;
    worksheet.write_number_only(4, 1, 1.22)?;

    workbook.save("future_function.xlsx")?;

    Ok(())
}
