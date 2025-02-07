// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates testing for a valid worksheet name.

use rust_xlsxwriter::{utility, XlsxError};

fn main() -> Result<(), XlsxError> {
    // This worksheet name is valid.
    let result = utility::check_sheet_name("2030-01-01")?;

    assert!(matches!(result, ()));

    // This worksheet name isn't valid due to the forward slashes.
    let result = utility::check_sheet_name("2030/01/01");

    assert!(matches!(
        result,
        Err(XlsxError::SheetnameContainsInvalidCharacter(_))
    ));

    Ok(())
}
