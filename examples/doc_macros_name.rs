// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates a simple example of adding a vba project
//! to an xlsm file.
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    workbook.add_vba_project("examples/vbaProject.bin")?;
    workbook.set_vba_name("MyWorkbook")?;

    let worksheet = workbook.add_worksheet();
    worksheet.set_vba_name("MySheet1")?;

    // Note the `.xlsm` extension.
    workbook.save("macros.xlsm")?;

    Ok(())
}
