// error - error values for the rust_xlsxwriter library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
use std::error::Error;
use std::fmt;

#[derive(Debug, PartialEq)]
/// Error values for the `rust_xlsxwriter` library.
pub enum XlsxError {
    /// Error returned when a row or column argument exceeds Excel's limits of
    /// 1,048,576 rows and 16,384 columns for a worksheet.
    RowColumnLimitError,

    /// Worksheet name cannot be blank.
    SheetnameCannotBeBlank,

    /// Worksheet name exceeds Excel's limit of 31 characters.
    SheetnameLengthExceeded,

    /// Worksheet name cannot contain invalid characters: `[ ] : * ? / \`
    SheetnameContainsInvalidCharacter,

    /// Worksheet name cannot start or end with an apostrophe.
    SheetnameStartsOrEndsWithApostrophe,

    /// String exceeds Excel's limit of 32,767 characters.
    MaxStringLengthExceeded,
}

impl Error for XlsxError {}

impl fmt::Display for XlsxError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            XlsxError::RowColumnLimitError => write!(
                f,
                "Row or column exceeds Excel's allowed limits (1,048,576 x 16,384)"
            ),
            XlsxError::SheetnameCannotBeBlank => write!(f, "Worksheet name cannot be blank. "),
            XlsxError::SheetnameLengthExceeded => {
                write!(f, "Worksheet name exceeds Excel's limit of 31 characters. ")
            }
            XlsxError::SheetnameContainsInvalidCharacter => write!(
                f,
                "Worksheet name cannot contain invalid characters: '[ ] : * ? / \\' "
            ),
            XlsxError::SheetnameStartsOrEndsWithApostrophe => {
                write!(f, "Worksheet name cannot start or end with an apostrophe. ")
            }
            XlsxError::MaxStringLengthExceeded => {
                write!(f, "String exceeds Excel's limit of 32,767 characters. ")
            }
        }
    }
}
