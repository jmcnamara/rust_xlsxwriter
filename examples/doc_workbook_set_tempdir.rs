// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting a custom directory for temporary
//! files when creating a file in "constant memory" mode.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Set a custom directory for temporary files before adding worksheets.
    workbook.set_tempdir(".")?;

    // Add a worksheet in "constant memory" mode.
    let worksheet = workbook.add_worksheet_with_constant_memory();
    worksheet.write(0, 0, "Hello")?;

    workbook.save("workbook.xlsx")?;

    Ok(())
}
