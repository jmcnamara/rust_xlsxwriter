// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a simple workbook from a Path,
//! with one unused worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let path = std::path::Path::new("workbook.xlsx");
    let mut workbook = Workbook::new_from_path(&path);

    _ = workbook.add_worksheet();

    workbook.close()?;

    Ok(())
}
