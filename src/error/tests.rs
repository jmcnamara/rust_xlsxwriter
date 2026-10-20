// Error unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

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
            XlsxError::IoError(Error::other("ERROR")).to_string(),
            "ERROR"
        );
        assert_eq!(
            XlsxError::ZipError(ZipError::FileNotFound).to_string(),
            "specified file not found in archive"
        );

        let error = XlsxError::TempFileError("ERROR".into(), Error::other("ERROR"));
        assert_eq!(error.to_string(), "Temporary file error: 'ERROR: ERROR'.");
        assert!(std::error::Error::source(&error).is_some());
        assert!(std::error::Error::source(&XlsxError::IoError(Error::other("ERROR"))).is_none());

        // A zip IO error forwards to the underlying IO error.
        let io_error = Error::new(ErrorKind::PermissionDenied, "ERROR");
        let error = XlsxError::ZipError(ZipError::Io(io_error));
        let source = std::error::Error::source(&error)
            .and_then(|source| source.downcast_ref::<Error>())
            .map(Error::kind);
        assert_eq!(Some(ErrorKind::PermissionDenied), source);

        assert!(std::error::Error::source(&XlsxError::ZipError(ZipError::FileNotFound)).is_none());
        assert!(std::error::Error::source(&XlsxError::RowColumnLimitError).is_none());

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
        Err(Error::other("ERROR"))
    }
}
