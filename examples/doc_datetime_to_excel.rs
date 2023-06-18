// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates the ExcelDateTime `to_excel()` method.

use rust_xlsxwriter::{ExcelDateTime, XlsxError};

fn main() -> Result<(), XlsxError> {
    let time = ExcelDateTime::from_hms(12, 0, 0)?;
    let date = ExcelDateTime::from_ymd(2000, 1, 1)?;

    assert_eq!(0.5, time.to_excel());
    assert_eq!(36526.0, date.to_excel());

    Ok(())
}
