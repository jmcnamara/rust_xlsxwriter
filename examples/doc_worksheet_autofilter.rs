// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting a simple autofilter in a
//! worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some header titles.
    worksheet.write_string_only(0, 0, "Region")?;
    worksheet.write_string_only(0, 1, "Count")?;

    // Write some test data.
    for row in 1..9 {
        worksheet.write_string_only(row as u32, 0, "East")?;
        worksheet.write_number_only(row as u32, 1, row * 100)?;
    }

    // Set the autofilter.
    worksheet.autofilter(0, 0, 8, 1)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
