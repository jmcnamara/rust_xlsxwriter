// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of adding a Custom UI XML file to modify the Ribbon in an Excel
//! file.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    workbook.add_custom_ui("examples/customUI14.xml")?;

    Ok(())
}
