// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates quoting worksheet names.

use rust_xlsxwriter::utility;

fn main() {
    // Doesn't need to be quoted.
    let result = utility::quote_sheet_name("Sheet1");
    assert_eq!(result, "Sheet1");

    // Spaces need to be quoted.
    let result = utility::quote_sheet_name("Sheet 1");
    assert_eq!(result, "'Sheet 1'");

    // Special characters need to be quoted.
    let result = utility::quote_sheet_name("Sheet-1");
    assert_eq!(result, "'Sheet-1'");

    // Single quotes need to be escaped with a quote.
    let result = utility::quote_sheet_name("Sheet'1");
    assert_eq!(result, "'Sheet''1'");

    // A1 style cell references don't need to be quoted.
    let result = utility::quote_sheet_name("A1");
    assert_eq!(result, "'A1'");

    // R1C1 style cell references need to be quoted.
    let result = utility::quote_sheet_name("RC1");
    assert_eq!(result, "'RC1'");
}
