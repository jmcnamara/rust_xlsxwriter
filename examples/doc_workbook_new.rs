// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook, with one
//! unused worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new("workbook.xlsx");

    _ = workbook.add_worksheet();

    workbook.close()?;

    Ok(())
}
