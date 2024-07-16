// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates a simple example of adding a vba project
//! to an xlsm file.
use rust_xlsxwriter::{Workbook, XlsxError};

#[allow(unused_variables)]
fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    workbook.add_vba_project_with_signature(
        "examples/vbaProject.bin",
        "examples/vbaProjectSignature.bin",
    )?;

    let worksheet = workbook.add_worksheet();

    // Note the `.xlsm` extension.
    workbook.save("macros.xlsm")?;

    Ok(())
}
