// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook using a Rust
//! Path reference.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let _worksheet = workbook.add_worksheet();

    let path = std::path::Path::new("workbook.xlsx");
    workbook.save(&path)?;

    Ok(())
}
