// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of adding macros to an `rust_xlsxwriter` file using a VBA macros
//! file extracted from an existing Excel xlsm file.
//!
//! The `vba_extract` utility (https://crates.io/crates/vba_extract) can be used
//! to extract the `vbaProject.bin` file.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add the VBA macro file.
    workbook.add_vba_project("examples/vbaProject.bin")?;

    // Add a worksheet and some text.
    let worksheet = workbook.add_worksheet();

    // Widen the first column for clarity.
    worksheet.set_column_width(0, 30)?;

    worksheet.write(2, 0, "Run macro say_hello()")?;

    // Save the file to disk. Note the `.xlsm` extension. This is required by
    // Excel or it raise a warning.
    workbook.save("macros.xlsm")?;

    Ok(())
}
