// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates testing for a valid worksheet name.

use rust_xlsxwriter::{utility, XlsxError};

fn main() -> Result<(), XlsxError> {
    // This worksheet name is valid.
    utility::check_sheet_name("2030-01-01")?;

    // This worksheet name isn't valid due to the forward slashes.
    utility::check_sheet_name("2030/01/01")?;

    Ok(())
}
