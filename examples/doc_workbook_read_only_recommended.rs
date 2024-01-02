// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook which opens
//! with a recommendation that the file should be opened in read only mode.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let _worksheet = workbook.add_worksheet();

    workbook.read_only_recommended();

    workbook.save("workbook.xlsx")?;

    Ok(())
}
