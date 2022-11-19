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

    /// Error when trying to retrieve a worksheet reference by index or by name.
    UnknownWorksheetNameOrIndex(String),

    /// A merge range cannot be a single cell in Excel.
    MergeRangeSingleCell,

    /// The merge range overlaps a previous merge range.
    MergeRangeOverlaps(String, String),

    /// URL string exceeds Excel's url of 2080 characters.
    MaxUrlLengthExceeded,

    /// Unknown url type. The URL/URIs supported by Excel are `http://`,
    /// `https://`, `ftp://`, `ftps://`, `mailto:`, `file://` and the
    /// pseudo-uri `internal:`:
    UnknownUrlType(String),

    /// Wrapper for a variety of [std::io::Error] errors such as file
    /// permissions when writing the xlsx file to disk. This can be caused by an
    /// non-existent parent directory or, commonly on Windows, if the file is
    /// already open in Excel.
    IoError(std::io::Error),

    /// Wrapper for a variety of [zip::result::ZipError] errors from
    /// [zip::ZipWriter]. These relate to errors arising from creating
    /// the xlsx file zip container.
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
                "First row or column in range is greater than last row or column."
            ),

            XlsxError::SheetnameCannotBeBlank => write!(f, "Worksheet name cannot be blank."),

            XlsxError::SheetnameLengthExceeded(name) => {
                write!(
                    f,
                    "Worksheet name \"{name}\" exceeds Excel's limit of 31 characters."
                )
            }

            XlsxError::SheetnameReused(name) => write!(
                f,
                "Worksheet name \"{name}\" has already been used in this workbook.",
            ),

            XlsxError::SheetnameContainsInvalidCharacter(name) => write!(
                f,
                "Worksheet name \"{name}\" cannot contain invalid characters: '[ ] : * ? / \\'.",
            ),

            XlsxError::SheetnameStartsOrEndsWithApostrophe(name) => {
                write!(
                    f,
                    "Worksheet name \"{name}\" cannot start or end with an apostrophe.",
                )
            }

            XlsxError::MaxStringLengthExceeded => {
                write!(f, "String exceeds Excel's limit of 32,767 characters.")
            }

            XlsxError::UnknownWorksheetNameOrIndex(name) => {
                write!(f, "Unknown Worksheet name or index \"{name}\".")
            }

            XlsxError::MergeRangeSingleCell => {
                write!(f, "A merge range cannot be a single cell in Excel.")
            }

            XlsxError::MergeRangeOverlaps(current, previous) => {
                write!(
                    f,
                    "Merge range {current} overlaps with previous merge range {previous}."
                )
            }

            XlsxError::MaxUrlLengthExceeded => {
                write!(f, "URL string exceeds Excel's limit of 2083 characters.")
            }

            XlsxError::UnknownUrlType(url) => {
                write!(f, "Unknown/unsupported url type: \"{url}\".")
            }

            XlsxError::IoError(error) => {
                write!(f, "{error}")
            }

            XlsxError::ZipError(error) => {
                write!(f, "{error}")
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

    use std::io::{Error, ErrorKind};

    use super::XlsxError;
    use pretty_assertions::assert_eq;
    use zip::result::ZipError;

    #[test]
    fn test_error_display() {
        let name = "ERROR";

        assert_eq!(
            XlsxError::RowColumnLimitError.to_string(),
            "Row or column exceeds Excel's allowed limits (1,048,576 x 16,384)."
        );
        assert_eq!(
            XlsxError::RowColumnOrderError.to_string(),
            "First row or column in range is greater than last row or column."
        );
        assert_eq!(
            XlsxError::SheetnameCannotBeBlank.to_string(),
            "Worksheet name cannot be blank."
        );
        assert_eq!(
            XlsxError::SheetnameLengthExceeded(name.to_string()).to_string(),
            "Worksheet name \"ERROR\" exceeds Excel's limit of 31 characters."
        );
        assert_eq!(
            XlsxError::SheetnameReused(name.to_string()).to_string(),
            "Worksheet name \"ERROR\" has already been used in this workbook."
        );
        assert_eq!(
            XlsxError::SheetnameContainsInvalidCharacter(name.to_string()).to_string(),
            "Worksheet name \"ERROR\" cannot contain invalid characters: '[ ] : * ? / \\'."
        );
        assert_eq!(
            XlsxError::SheetnameStartsOrEndsWithApostrophe(name.to_string()).to_string(),
            "Worksheet name \"ERROR\" cannot start or end with an apostrophe."
        );
        assert_eq!(
            XlsxError::MaxStringLengthExceeded.to_string(),
            "String exceeds Excel's limit of 32,767 characters."
        );
        assert_eq!(
            XlsxError::UnknownWorksheetNameOrIndex(name.to_string()).to_string(),
            "Unknown Worksheet name or index \"ERROR\"."
        );
        assert_eq!(
            XlsxError::MergeRangeSingleCell.to_string(),
            "A merge range cannot be a single cell in Excel."
        );
        assert_eq!(
            XlsxError::MergeRangeOverlaps(name.to_string(), name.to_string()).to_string(),
            "Merge range ERROR overlaps with previous merge range ERROR."
        );

        assert_eq!(
            XlsxError::IoError(Error::new(ErrorKind::Other, "ERROR")).to_string(),
            "ERROR"
        );
        assert_eq!(
            XlsxError::ZipError(ZipError::FileNotFound).to_string(),
            "specified file not found in archive"
        );

        let result = catch_zip_error();
        assert!(matches!(result, Err(XlsxError::ZipError(_))));

        let result = catch_io_error();
        assert!(matches!(result, Err(XlsxError::IoError(_))));

        assert_eq!(
            format!("{:?}", XlsxError::RowColumnLimitError),
            "RowColumnLimitError"
        );
    }

    fn catch_zip_error() -> Result<(), XlsxError> {
        throw_zip_error()?;
        Ok(())
    }

    fn throw_zip_error() -> Result<(), ZipError> {
        Err(ZipError::FileNotFound)
    }

    fn catch_io_error() -> Result<(), XlsxError> {
        throw_io_error()?;
        Ok(())
    }

    fn throw_io_error() -> Result<(), std::io::Error> {
        Err(Error::new(ErrorKind::Other, "ERROR"))
    }
}
