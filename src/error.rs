// error - error values for the rust_xlsxwriter library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
use std::error::Error;
use std::fmt;

#[derive(Debug)]
/// Error values for the `rust_xlsxwriter` library.
pub enum XlsxError {
    /// Error returned when a row or column argument exceeds Excel's limits of
    /// 1,048,576 rows and 16,384 columns for a worksheet.
    RowColumnLimitError,

    /// First row or column is greater than last row or column in a range
    /// specification, i.e., the order is reversed.
    RowColumnOrderError,

    /// Worksheet name cannot be blank.
    SheetnameCannotBeBlank,

    /// Worksheet name exceeds Excel's limit of 31 characters.
    SheetnameLengthExceeded(String),

    /// Worksheet name is already in use in the workbook.
    SheetnameReused(String),

    /// Worksheet name cannot contain invalid characters: `[ ] : * ? / \`
    SheetnameContainsInvalidCharacter(String),

    /// Worksheet name cannot start or end with an apostrophe.
    SheetnameStartsOrEndsWithApostrophe(String),

    /// String exceeds Excel's limit of 32,767 characters.
    MaxStringLengthExceeded,

    /// Cannot close() file that is already closed.
    FileReClosedError,

    /// Wrapper for a variety of [std::io::Error] errors such as file
    /// permissions when writing the xlsx file to disk. This can be caused by an
    /// non-existent parent directory or, commonly on Windows, if the file is
    /// already open in Excel.
    IoError(std::io::Error),

    /// Wrapper for a variety of [zip::result::ZipError] errors from
    /// [zip::ZipWriter]. These relate to errors arising from creating
    /// the XLSX file zip container.
    ZipError(zip::result::ZipError),
}

impl Error for XlsxError {}

impl fmt::Display for XlsxError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            XlsxError::RowColumnLimitError => write!(
                f,
                "Row or column exceeds Excel's allowed limits (1,048,576 x 16,384)."
            ),

            XlsxError::RowColumnOrderError => write!(
                f,
                "First row or column in range is greater than last row or column. "
            ),

            XlsxError::SheetnameCannotBeBlank => write!(f, "Worksheet name cannot be blank."),

            XlsxError::SheetnameLengthExceeded(name) => {
                write!(
                    f,
                    "Worksheet name \"{}\" exceeds Excel's limit of 31 characters.",
                    name
                )
            }

            XlsxError::SheetnameReused(name) => write!(
                f,
                "Worksheet name \"{}\" has already been used in this workbook.",
                name
            ),

            XlsxError::SheetnameContainsInvalidCharacter(name) => write!(
                f,
                "Worksheet name \"{}\" cannot contain invalid characters: '[ ] : * ? / \\'.",
                name
            ),

            XlsxError::SheetnameStartsOrEndsWithApostrophe(name) => {
                write!(
                    f,
                    "Worksheet name \"{}\" cannot start or end with an apostrophe.",
                    name
                )
            }

            XlsxError::MaxStringLengthExceeded => {
                write!(f, "String exceeds Excel's limit of 32,767 characters.")
            }

            XlsxError::FileReClosedError => write!(f, "File has already been closed."),

            XlsxError::IoError(e) => {
                write!(f, "{}", e)
            }

            XlsxError::ZipError(e) => {
                write!(f, "{}", e)
            }
        }
    }
}

// Convert errors from ZipWriter.
impl From<zip::result::ZipError> for XlsxError {
    fn from(e: zip::result::ZipError) -> XlsxError {
        XlsxError::ZipError(e)
    }
}

// Convert IO errors that arise directly or from ZipWriter.
impl From<std::io::Error> for XlsxError {
    fn from(e: std::io::Error) -> XlsxError {
        XlsxError::IoError(e)
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::XlsxError;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_error_display() {
        let name = "Foo";

        assert_eq!(
            XlsxError::RowColumnLimitError.to_string(),
            "Row or column exceeds Excel's allowed limits (1,048,576 x 16,384)."
        );
        assert_eq!(
            XlsxError::SheetnameCannotBeBlank.to_string(),
            "Worksheet name cannot be blank."
        );
        assert_eq!(
            XlsxError::SheetnameLengthExceeded(name.to_string()).to_string(),
            "Worksheet name \"Foo\" exceeds Excel's limit of 31 characters."
        );
        assert_eq!(
            XlsxError::SheetnameContainsInvalidCharacter(name.to_string()).to_string(),
            "Worksheet name \"Foo\" cannot contain invalid characters: '[ ] : * ? / \\'."
        );
        assert_eq!(
            XlsxError::SheetnameStartsOrEndsWithApostrophe(name.to_string()).to_string(),
            "Worksheet name \"Foo\" cannot start or end with an apostrophe."
        );
        assert_eq!(
            XlsxError::MaxStringLengthExceeded.to_string(),
            "String exceeds Excel's limit of 32,767 characters."
        );
    }
}
