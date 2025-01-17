// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the worksheet properties to be
//! protected in a protected worksheet. In this case we protect the overall
//! worksheet but allow columns and rows to be inserted.

use rust_xlsxwriter::{ProtectionOptions, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Set some of the options and use the defaults for everything else.
    let options = ProtectionOptions {
        insert_columns: true,
        insert_rows: true,
        ..ProtectionOptions::default()
    };

    // Set the protection options.
    worksheet.protect_with_options(&options);

    worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
