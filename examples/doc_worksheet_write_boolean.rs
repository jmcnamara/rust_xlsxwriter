// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing formatted boolean values to a
//! worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let bold = Format::new().set_bold();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    worksheet.write_boolean(0, 0, true, &bold)?;
    worksheet.write_boolean(1, 0, false, &bold)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
