// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing boolean values to a worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.write_boolean(0, 0, true)?;
    worksheet.write_boolean(1, 0, false)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
