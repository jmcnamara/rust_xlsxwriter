// error - error values for the `rust_xlsxwriter` library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::error::Error;
use std::fmt;

#[cfg(feature = "polars")]
use polars::prelude::polars_err;

#[cfg(feature = "polars")]
use polars::prelude::PolarsError;

#[cfg(feature = "serde")]
use serde::de;

#[cfg(feature = "serde")]
use serde::ser;

#[derive(Debug)]
/// The `XlsxError` enum defines the error values for the `rust_xlsxwriter`
/// library.
pub enum XlsxError {
    /// A general parameter error that is raised when a parameter conflicts with
    /// an Excel limit or syntax. The nature of the error is in the error string.
    ParameterError(String),

    /// Row or column argument exceeds Excel's limits of 1,048,576 rows and
    /// 16,384 columns for a worksheet.
    RowColumnLimitError,

    /// First row or column is greater than last row or column in a range
    /// specification, i.e., the order is reversed.
    RowColumnOrderError,

    /// Worksheet name cannot be blank.
    SheetnameCannotBeBlank(String),

    /// Worksheet name exceeds Excel's limit of 31 characters.
    SheetnameLengthExceeded(String),

    /// Worksheet name is already in use in the workbook.
    SheetnameReused(String),

    /// Worksheet name cannot contain any of the following invalid characters: `[ ] : * ? / \`
    SheetnameContainsInvalidCharacter(String),

    /// Worksheet name cannot start or end with an apostrophe.
    SheetnameStartsOrEndsWithApostrophe(String),

    /// String exceeds Excel's limit of 32,767 characters.
    MaxStringLengthExceeded,

    /// Error when trying to retrieve a worksheet reference by index or by name.
    UnknownWorksheetNameOrIndex(String),

    /// A merge range cannot be a single cell in Excel.
    MergeRangeSingleCell,

    /// The merge range overlaps a previous merge range. This is a strictly
    /// prohibited by Excel.
    MergeRangeOverlaps(String, String),

    /// URL string exceeds Excel's url of 2080 characters.
    MaxUrlLengthExceeded,

    /// Unknown url type. The URL/URIs supported by Excel are `http://`,
    /// `https://`, `ftp://`, `ftps://`, `mailto:`, `file://` and the
    /// pseudo-uri `internal:`.
    UnknownUrlType(String),

    /// Unknown image type. The supported image formats are PNG, JPG, GIF and
    /// BMP. See [`Image`](crate::Image) for details.
    UnknownImageType,

    /// Image has zero width or height, or the dimensions couldn't be read.
    ImageDimensionError,

    /// A general error that is raised when a chart parameter is incorrect, or a
    /// chart is configured incorrectly.
    ChartError(String),

    /// A general error that is raised when a sparkline parameter is incorrect,
    /// or a sparkline is configured incorrectly.
    SparklineError(String),

    /// A general error when one of the parameters supplied to a
    /// [`ExcelDateTime`](crate::ExcelDateTime) method is outside Excel's
    /// allowable ranges.
    ///
    /// Excel restricts dates to the range 1899-12-31 to 9999-12-31. For hours
    /// the range is generally 0-24 although larger ranges can be used to
    /// indicate durations. Minutes should be in the range 0-60 and seconds
    /// should be in the range 0.0-59.999. Excel only supports millisecond
    /// resolution.
    DateTimeRangeError(String),

    /// A parsing error when trying to convert a string into an
    /// [`ExcelDateTime`](crate::ExcelDateTime).
    ///
    /// The allowable date/time formats supported by
    /// [`ExcelDateTime::parse_from_str()`](crate::ExcelDateTime::parse_from_str)
    /// are:
    ///
    /// ```text
    /// Dates:
    ///     yyyy-mm-dd
    ///
    /// Times:
    ///     hh:mm
    ///     hh:mm:ss
    ///     hh:mm:ss.sss
    ///
    /// DateTimes:
    ///     yyyy-mm-ddThh:mm:ss
    ///     yyyy-mm-dd hh:mm:ss
    /// ```
    ///
    /// The time part of `DateTimes` can contain optional or fractional seconds
    /// like the time examples. Timezone information is not supported by Excel
    /// and ignored in the parsing.
    ///
    DateTimeParseError(String),

    /// The table range overlaps a previous table range. This is a strictly
    /// prohibited by Excel.
    TableRangeOverlaps(String, String),

    /// A general error that is raised when a table parameter is incorrect, or a
    /// table is configured incorrectly.
    TableError(String),

    /// Table name is already in use in the workbook.
    TableNameReused(String),

    /// A general error that is raised when a conditional format parameter is
    /// incorrect or missing.
    ConditionalFormatError(String),

    /// A general error that is raised when a data validation parameter is
    /// incorrect or missing.
    DataValidationError(String),

    /// A general error raised when a VBA name doesn't meet Excel's criteria as
    /// defined by the following rules:
    ///
    /// - The name must be less than 32 characters.
    /// - The name can only contain word characters: letters, numbers and
    ///   underscores.
    /// - The name must start with a letter.
    /// - The name cannot be blank.
    ///
    VbaNameError(String),

    /// A customizable error that can be used by third parties to raise errors
    /// or as a conversion target for other Error types.
    CustomError(String),

    /// Wrapper for a variety of [`std::io::Error`] errors such as file
    /// permissions when writing the xlsx file to disk. This can be caused by a
    /// non-existent parent directory or, commonly on Windows, if the file is
    /// already open in Excel.
    IoError(std::io::Error),

    /// Wrapper for a variety of [`zip::result::ZipError`] errors from
    /// [`zip::ZipWriter`]. These relate to errors arising from creating
    /// the xlsx file zip container.
    ZipError(zip::result::ZipError),

    /// A general error that is raised when serializing data via the Serde
    /// serializer. This requires the `serde` feature to be enabled.
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    SerdeError(String),

    /// Wrapper for a variety of [polars::prelude::PolarsError] errors. This is
    /// mainly used by the `polars_excel_writer` crate but it can also be useful
    /// for code that uses `polars` functions in an `XlsxError` error scope.
    /// This requires the `polars` feature to be enabled.
    #[cfg(feature = "polars")]
    #[cfg_attr(docsrs, doc(cfg(feature = "polars")))]
    PolarsError(PolarsError),
}

impl Error for XlsxError {}

impl fmt::Display for XlsxError {
    #[allow(clippy::too_many_lines)]
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            XlsxError::ParameterError(error) => {
                write!(f, "Parameter error: '{error}'.")
            }

            XlsxError::RowColumnLimitError => write!(
                f,
                "Row or column exceeds Excel's allowed limits (1,048,576 x 16,384)."
            ),

            XlsxError::RowColumnOrderError => write!(
                f,
                "First row or column in range is greater than last row or column."
            ),

            XlsxError::SheetnameCannotBeBlank(name) => {
                write!(f, "Worksheet name '{name}' cannot be blank.")
            }

            XlsxError::SheetnameLengthExceeded(name) => {
                write!(
                    f,
                    "Worksheet name '{name}' exceeds Excel's limit of 31 characters."
                )
            }

            XlsxError::SheetnameReused(name) => write!(
                f,
                "Worksheet name '{name}' has already been used in this workbook.",
            ),

            XlsxError::SheetnameContainsInvalidCharacter(name) => write!(
                f,
                "Worksheet name '{name}' cannot contain invalid characters: '[ ] : * ? / \\'.",
            ),

            XlsxError::SheetnameStartsOrEndsWithApostrophe(name) => {
                write!(
                    f,
                    "Worksheet name '{name}' cannot start or end with an apostrophe.",
                )
            }

            XlsxError::MaxStringLengthExceeded => {
                write!(f, "String exceeds Excel's limit of 32,767 characters.")
            }

            XlsxError::UnknownWorksheetNameOrIndex(name) => {
                write!(f, "Unknown Worksheet name or index '{name}'.")
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

            XlsxError::TableRangeOverlaps(current, previous) => {
                write!(
                    f,
                    "Table range {current} overlaps with previous table range {previous}."
                )
            }

            XlsxError::MaxUrlLengthExceeded => {
                write!(f, "URL string exceeds Excel's limit of 2083 characters.")
            }

            XlsxError::UnknownUrlType(url) => {
                write!(f, "Unknown/unsupported url type: '{url}'.")
            }

            XlsxError::UnknownImageType => {
                write!(f, "Unknown image type.")
            }

            XlsxError::ImageDimensionError => {
                write!(f, "Image with or height couldn't be read from file.")
            }

            XlsxError::ChartError(error) => {
                write!(f, "Chart error: '{error}'.")
            }

            XlsxError::SparklineError(error) => {
                write!(f, "Sparkline error: '{error}'.")
            }

            XlsxError::DateTimeRangeError(error) => {
                write!(f, "Date range error: '{error}'")
            }

            XlsxError::DateTimeParseError(error) => {
                write!(f, "Date parse error: '{error}'")
            }

            XlsxError::TableError(error) => {
                write!(f, "Table error: '{error}'.")
            }

            XlsxError::TableNameReused(name) => {
                write!(
                    f,
                    "Table name '{name}' has already been used in this workbook.",
                )
            }

            XlsxError::ConditionalFormatError(error) => {
                write!(f, "Conditional format error: '{error}'.")
            }

            XlsxError::DataValidationError(error) => {
                write!(f, "Data validation error: '{error}'.")
            }

            XlsxError::VbaNameError(error) => {
                write!(f, "VBA name error: '{error}'.")
            }

            XlsxError::CustomError(error) => {
                write!(f, "{error}")
            }

            XlsxError::IoError(error) => {
                write!(f, "{error}")
            }

            XlsxError::ZipError(error) => {
                write!(f, "{error}")
            }

            #[cfg(feature = "serde")]
            XlsxError::SerdeError(error) => {
                write!(f, "Serialization error: '{error}'.")
            }

            #[cfg(feature = "polars")]
            XlsxError::PolarsError(error) => {
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

// Convert from Polars to Polars errors to allow easier interoperability.
#[cfg(feature = "polars")]
impl From<PolarsError> for XlsxError {
    fn from(e: PolarsError) -> XlsxError {
        XlsxError::PolarsError(e)
    }
}

// Convert from XlsxError to Polars errors to allow easier interoperability.
#[cfg(feature = "polars")]
impl From<XlsxError> for PolarsError {
    fn from(e: XlsxError) -> PolarsError {
        polars_err!(ComputeError: "rust_xlsxwriter error: '{}'", e)
    }
}

// Convert from XlsxError to JsValue errors to allow easier interoperability.
#[cfg(all(
    feature = "wasm",
    target_arch = "wasm32",
    not(any(target_os = "emscripten", target_os = "wasi"))
))]
impl From<XlsxError> for wasm_bindgen::JsValue {
    fn from(e: XlsxError) -> wasm_bindgen::JsValue {
        let error = e.to_string();
        wasm_bindgen::JsValue::from_str(&error)
    }
}

/// Implementation of the `serde::de::Error` and `serde::ser::Error` Traits to
/// allow the use of a single error type for deserialization/serialization and
/// `rust_xlsxwriter` errors.
#[cfg(feature = "serde")]
impl ser::Error for XlsxError {
    fn custom<T: fmt::Display>(msg: T) -> Self {
        XlsxError::SerdeError(msg.to_string())
    }
}

#[cfg(feature = "serde")]
impl de::Error for XlsxError {
    fn custom<T: fmt::Display>(msg: T) -> Self {
        XlsxError::SerdeError(msg.to_string())
    }
}
