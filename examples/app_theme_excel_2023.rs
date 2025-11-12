// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! Example of changing the default theme for a workbook using the
//! `rust_xlsxwriter` library. The example uses the Excel 2023 Office/Aptos
//! theme.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Use the Excel 2023 Office/Aptos theme in the workbook.
    workbook.use_excel_2023_theme()?;

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some text to demonstrate the changed theme.
    worksheet.write(0, 0, "Hello")?;

    // Save the workbook to disk.
    workbook.save("theme_excel_2023.xlsx")?;

    Ok(())
}
