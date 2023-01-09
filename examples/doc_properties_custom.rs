// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of setting custom/user defined workbook document properties.

use rust_xlsxwriter::{Properties, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let properties = Properties::new()
        .set_custom_text("Checked by", "Admin")
        .set_custom_boolean("Cross check", true)
        .set_custom_text("Department", "Finance")
        .set_custom_number_i32("Document number", 55301);

    workbook.set_properties(&properties);

    workbook.save("properties.xlsx")?;

    Ok(())
}
