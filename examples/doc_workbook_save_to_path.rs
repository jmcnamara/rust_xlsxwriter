// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook using a rust
//! Path reference.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let path = std::path::Path::new("workbook.xlsx");
    let mut workbook = Workbook::new();

    _ = workbook.add_worksheet();

    workbook.save(&path)?;

    Ok(())
}
