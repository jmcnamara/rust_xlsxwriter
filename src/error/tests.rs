// Error unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod error_tests {

    use std::io::{Error, ErrorKind};

    use crate::XlsxError;
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
            XlsxError::SheetnameCannotBeBlank(name.to_string()).to_string(),
            "Worksheet name 'ERROR' cannot be blank."
        );
        assert_eq!(
            XlsxError::SheetnameLengthExceeded(name.to_string()).to_string(),
            "Worksheet name 'ERROR' exceeds Excel's limit of 31 characters."
        );
        assert_eq!(
            XlsxError::SheetnameReused(name.to_string()).to_string(),
            "Worksheet name 'ERROR' has already been used in this workbook."
        );
        assert_eq!(
            XlsxError::SheetnameContainsInvalidCharacter(name.to_string()).to_string(),
            "Worksheet name 'ERROR' cannot contain invalid characters: '[ ] : * ? / \\'."
        );
        assert_eq!(
            XlsxError::SheetnameStartsOrEndsWithApostrophe(name.to_string()).to_string(),
            "Worksheet name 'ERROR' cannot start or end with an apostrophe."
        );
        assert_eq!(
            XlsxError::MaxStringLengthExceeded.to_string(),
            "String exceeds Excel's limit of 32,767 characters."
        );
        assert_eq!(
            XlsxError::UnknownWorksheetNameOrIndex(name.to_string()).to_string(),
            "Unknown Worksheet name or index 'ERROR'."
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
