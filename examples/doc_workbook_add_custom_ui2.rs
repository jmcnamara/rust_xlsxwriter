// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of adding both the Excel 2007 and the Excel 2010+ format Custom
//! UI XML files to a workbook.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add an Excel 2007 format ribbon.
    workbook.add_custom_ui("examples/customUI.xml")?;

    // Add an Excel 2010+ format ribbon.
    workbook.add_custom_ui("examples/customUI14.xml")?;

    Ok(())
}
