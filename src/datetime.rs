// excel_datetime - A module for handling Excel dates and times.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
mod tests;

#[cfg(feature = "serde")]
use serde::{Deserialize, Deserializer, Serialize, Serializer};

#[cfg(feature = "chrono")]
use chrono::{Datelike, NaiveDate, NaiveDateTime, NaiveTime};

#[cfg(not(all(
    feature = "wasm",
    target_arch = "wasm32",
    not(any(target_os = "emscripten", target_os = "wasi"))
)))]
use std::time::SystemTime;

use crate::XlsxError;

const DAY_SECONDS: u64 = 24 * 60 * 60;
const HOUR_SECONDS: u64 = 60 * 60;
const MINUTE_SECONDS: u64 = 60;
const YEAR_DAYS: u64 = 365;
const YEAR_DAYS_4: u64 = YEAR_DAYS * 4 + 1;
const YEAR_DAYS_100: u64 = YEAR_DAYS * 100 + 25;
const YEAR_DAYS_400: u64 = YEAR_DAYS * 400 + 97;
const UNIX_EPOCH_PLUS_400: i64 = 12_622_780_800;

/// The `ExcelDateTime` struct is used to represent an Excel date and/or time.
///
/// The `rust_xlsxwriter` library supports two ways of converting dates and
/// times to Excel dates and times. The first is the inbuilt [`ExcelDateTime`]
/// which has a limited but workable set of conversion methods and which only
/// targets Excel specific dates and times. The second is via the external
/// [`Chrono`] library which has a comprehensive sets of types and functions for
/// dealing with dates and times.
///
/// [`Chrono`]: https://docs.rs/chrono/latest/chrono
///
/// Here is an example using `ExcelDateTime` to write some dates and times:
///
/// ```
/// # // This code is available in examples/doc_datetime_intro.rs
/// #
/// use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Create some formats to use with the datetimes below.
///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
///     let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
///     let format4 = Format::new().set_num_format("ddd dd mmm yyyy hh:mm");
///     let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy hh:mm");
///
///     // Set the column width for clarity.
///     worksheet.set_column_width(0, 30)?;
///
///     // Create a datetime object.
///     let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 0)?;
///
///     // Write the datetime with different Excel formats.
///     worksheet.write_with_format(0, 0, &datetime, &format1)?;
///     worksheet.write_with_format(1, 0, &datetime, &format2)?;
///     worksheet.write_with_format(2, 0, &datetime, &format3)?;
///     worksheet.write_with_format(3, 0, &datetime, &format4)?;
///     worksheet.write_with_format(4, 0, &datetime, &format5)?;
///
///     workbook.save("datetime.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/datetime_intro.png">
///
/// ## Datetimes in Excel
///
/// Datetimes in Excel are serial dates with days counted from an epoch (usually
/// 1900-01-01) and where the time is a percentage/decimal of the milliseconds
/// in the day. Both the date and time are stored in the same f64 value. For
/// example, 2023/01/01 12:00:00 is stored as 44927.5.
///
/// Datetimes in Excel must also be formatted with a number format like
/// `"yyyy/mm/dd hh:mm"` or otherwise they will appear as a number (which
/// technically they are).
///
/// Excel doesn't use timezones or try to convert or encode timezone information
/// in any way so they aren't supported by `rust_xlsxwriter`.
///
/// Excel can also save dates in a text ISO 8601 format when the file is saved
/// using the "Strict Open XML Spreadsheet" option in the "Save" dialog. However
/// this is rarely used in practice and isn't supported by `rust_xlsxwriter`.
///
/// ## Chrono vs. native `ExcelDateTime`
///
/// The `rust_xlsxwriter` native `ExcelDateTime` provided most of the
/// functionality that you will need to work with Excel dates and times.
///
/// For anything more advanced you can use the Naive Date/Time variants of
/// [`Chrono`], particularly if you are interacting with code that already uses
/// `Chrono`.
///
/// All date/time APIs in `rust_xlsxwriter` support both options and the
/// `ExcelDateTime` method names are similar to `Chrono` method names to allow
/// easier portability between the two.
///
/// In order to use [`Chrono`] with `rust_xlsxwriter` APIs you must enable the
/// optional `chrono` feature when adding `rust_xlsxwriter` to your
/// `Cargo.toml`.
///
/// ```bash
/// cargo add rust_xlsxwriter -F chrono
/// ```
///
/// [`Chrono`]: https://docs.rs/chrono/latest/chrono
///
#[derive(Clone)]
pub struct ExcelDateTime {
    year: u16,
    month: u8,
    day: u8,
    hour: u16,
    min: u8,
    sec: f64,
    is_1904_date: bool,
    serial_datetime: Option<f64>,
    datetime_type: ExcelDateTimeType,
}

impl ExcelDateTime {
    /// Create a `ExcelDateTime` instance from a string reference.
    ///
    /// This method provides simple conversions from strings representing dates,
    /// times and datetimes in approximate ISO 8601 format to `ExcelDateTime`
    /// instances.
    ///
    /// The allowable formats are:
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
    ///
    /// ```
    ///
    /// Notes:
    ///
    /// 1. The time portion of `DateTimes` can contain optional or fractional
    ///    seconds like the `Times` examples.
    /// 2. Leading or trailing whitespace or text that isn't part of the
    ///    date/times is ignored. For example a trailing `Z` in the datetime is
    ///    ignored.
    /// 3. Timezones aren't handled by Excel and are ignored in the input
    ///    string.
    /// 4. The `parse_to_str()` method is deliberately simple and limited. It
    ///    doesn't implement anything like a `strftime()` method. For more
    ///    comprehensive date parsing you should use the [`Chrono`] library.
    ///
    /// [`Chrono`]: https://docs.rs/chrono/latest/chrono
    ///
    /// # Parameters
    ///
    /// `datetime` - A string representing a date, time or datetime in the
    /// formats shown above.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    /// - [`XlsxError::DateTimeParseError`] - The input string couldn't be parsed
    ///   into a date/time.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing datetimes parsed from
    /// strings.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_parse_from_str.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("hh:mm:ss");
    ///     let format2 = Format::new().set_num_format("yyyy-mm-dd");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create different datetime objects.
    ///     let datetime1 = ExcelDateTime::parse_from_str("12:30")?;
    ///     let datetime2 = ExcelDateTime::parse_from_str("12:30:45")?;
    ///     let datetime3 = ExcelDateTime::parse_from_str("12:30:45.5")?;
    ///     let datetime4 = ExcelDateTime::parse_from_str("2023-01-31")?;
    ///     let datetime5 = ExcelDateTime::parse_from_str("2023-01-31 12:30:45")?;
    ///     let datetime6 = ExcelDateTime::parse_from_str("2023-01-31T12:30:45Z")?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_with_format(0, 0, &datetime1, &format1)?;
    ///     worksheet.write_with_format(1, 0, &datetime2, &format1)?;
    ///     worksheet.write_with_format(2, 0, &datetime3, &format1)?;
    ///     worksheet.write_with_format(3, 0, &datetime4, &format2)?;
    ///     worksheet.write_with_format(4, 0, &datetime5, &format3)?;
    ///     worksheet.write_with_format(5, 0, &datetime6, &format3)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/datetime_parse_from_str.png">
    ///
    #[allow(clippy::get_first)]
    pub fn parse_from_str(datetime: &str) -> Result<ExcelDateTime, XlsxError> {
        let date_parts: Vec<&str> = datetime.trim().split(&['-', 'T', 'Z', ' ', ':']).collect();

        // Match date and optional time.
        if datetime.contains('-') {
            let year = match date_parts.get(0) {
                Some(token) => token.parse::<u16>().unwrap_or_default(),
                None => 0,
            };

            let month = match date_parts.get(1) {
                Some(token) => token.parse::<u8>().unwrap_or_default(),
                None => 0,
            };

            let day = match date_parts.get(2) {
                Some(token) => token.parse::<u8>().unwrap_or_default(),
                None => 0,
            };

            let hour = match date_parts.get(3) {
                Some(token) => token.parse::<u16>().unwrap_or_default(),
                None => 0,
            };

            let min = match date_parts.get(4) {
                Some(token) => token.parse::<u8>().unwrap_or_default(),
                None => 0,
            };

            let sec = match date_parts.get(5) {
                Some(token) => token.parse::<f64>().unwrap_or_default(),
                None => 0.0,
            };

            Ok(ExcelDateTime::from_ymd(year, month, day)?.and_hms(hour, min, sec)?)
        }
        // Match time only.
        else if datetime.contains(':') {
            let hour = match date_parts.get(0) {
                Some(token) => token.parse::<u16>().unwrap_or_default(),
                None => 0,
            };

            let min = match date_parts.get(1) {
                Some(token) => token.parse::<u8>().unwrap_or_default(),
                None => 0,
            };

            let sec = match date_parts.get(2) {
                Some(token) => token.parse::<f64>().unwrap_or_default(),
                None => 0.0,
            };

            Ok(ExcelDateTime::from_hms(hour, min, sec)?)
        }
        // No match.
        else {
            Err(XlsxError::DateTimeParseError(datetime.to_string()))
        }
    }

    /// Create a `ExcelDateTime` instance from years, months and days.
    ///
    /// # Parameters
    ///
    /// - `year`: Integer year in range 1900-9999.
    /// - `month`: Integer month in the range 1-12.
    /// - `day`: Integer day in the range 1-31 (depending on year/month).
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges. Excel dates must be in the
    ///   range 1900-01-01 to 9999-12-31.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted dates in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_and_hms.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    ///     let format4 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.0");
    ///     let format5 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.000");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 45.195)?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_with_format(0, 0, &datetime, &format1)?;
    ///     worksheet.write_with_format(1, 0, &datetime, &format2)?;
    ///     worksheet.write_with_format(2, 0, &datetime, &format3)?;
    ///     worksheet.write_with_format(3, 0, &datetime, &format4)?;
    ///     worksheet.write_with_format(4, 0, &datetime, &format5)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/datetime_from_ymd.png">
    ///
    pub fn from_ymd(year: u16, month: u8, day: u8) -> Result<ExcelDateTime, XlsxError> {
        if let Some(err) = Self::validate_ymd(year, month, day).err() {
            return Err(err);
        }

        let dt = ExcelDateTime {
            year,
            month,
            day,
            datetime_type: ExcelDateTimeType::DateOnly,
            ..ExcelDateTime::default()
        };

        Ok(dt)
    }

    /// Create a `ExcelDateTime` instance from hours, minutes and seconds.
    ///
    /// # Parameters
    ///
    /// - `hour`: Integer hour. Generally in the range 0-23 but can be greater
    ///   than 24 for time durations.
    /// - `min`: Integer minutes in the range 0-59.
    /// - `sec`: Integer or float seconds in the range 0-59.999. Excel only
    ///   supports millisecond precision.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted times in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_from_hms.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("hh:mm");
    ///     let format2 = Format::new().set_num_format("hh:mm:ss");
    ///     let format3 = Format::new().set_num_format("hh:mm:ss.0");
    ///     let format4 = Format::new().set_num_format("hh:mm:ss.000");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_hms(12, 30, 45.5)?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_with_format(0, 0, &datetime, &format1)?;
    ///     worksheet.write_with_format(1, 0, &datetime, &format2)?;
    ///     worksheet.write_with_format(2, 0, &datetime, &format3)?;
    ///     worksheet.write_with_format(3, 0, &datetime, &format4)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/datetime_from_hms.png">
    ///
    ///
    pub fn from_hms(hour: u16, min: u8, sec: impl Into<f64>) -> Result<ExcelDateTime, XlsxError> {
        ExcelDateTime::default().and_hms(hour, min, sec)
    }

    /// Create a `ExcelDateTime` instance from hours, minutes, seconds and
    /// milliseconds.
    ///
    /// # Parameters
    ///
    /// - `hour`: Integer hour. Generally in the range 0-23 but can be greater
    ///   than 24 for time durations.
    /// - `min`: Integer minutes in the range 0-59.
    /// - `sec`: Integer seconds in the range 0-59.
    /// - `milli`: Integer milliseconds in the range 0-999. Excel only supports
    ///   millisecond precision.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted times in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_from_hms_milli.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("hh:mm");
    ///     let format2 = Format::new().set_num_format("hh:mm:ss");
    ///     let format3 = Format::new().set_num_format("hh:mm:ss.0");
    ///     let format4 = Format::new().set_num_format("hh:mm:ss.000");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_hms_milli(12, 30, 45, 123)?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_with_format(0, 0, &datetime, &format1)?;
    ///     worksheet.write_with_format(1, 0, &datetime, &format2)?;
    ///     worksheet.write_with_format(2, 0, &datetime, &format3)?;
    ///     worksheet.write_with_format(3, 0, &datetime, &format4)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/datetime_from_hms_milli.png">
    ///
    ///
    pub fn from_hms_milli(
        hour: u16,
        min: u8,
        sec: u8,
        milli: u16,
    ) -> Result<ExcelDateTime, XlsxError> {
        ExcelDateTime::default().and_hms_milli(hour, min, sec, milli)
    }

    /// Adds to a `ExcelDateTime` date instance with hours, minutes and seconds.
    ///
    /// Adds time to a existing `ExcelDateTime` date instance or creates a new
    /// one if required.
    ///
    /// # Parameters
    ///
    /// - `hour`: Integer hours in the range 0-23.
    /// - `min`: Integer minutes in the range 0-59.
    /// - `sec`: Integer or float seconds in the range 0-59.999. Excel only
    ///   supports millisecond precision.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted datetimes in an
    /// Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_and_hms.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    ///     let format4 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.0");
    ///     let format5 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.000");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 45.195)?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_with_format(0, 0, &datetime, &format1)?;
    ///     worksheet.write_with_format(1, 0, &datetime, &format2)?;
    ///     worksheet.write_with_format(2, 0, &datetime, &format3)?;
    ///     worksheet.write_with_format(3, 0, &datetime, &format4)?;
    ///     worksheet.write_with_format(4, 0, &datetime, &format5)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/datetime_and_hms.png">
    ///
    pub fn and_hms(
        mut self,
        hour: u16,
        min: u8,
        sec: impl Into<f64>,
    ) -> Result<ExcelDateTime, XlsxError> {
        let sec = sec.into();

        if let Some(err) = Self::validate_hms(min, sec).err() {
            return Err(err);
        }

        let date_time_type = if self.datetime_type == ExcelDateTimeType::DateOnly {
            ExcelDateTimeType::DateAndTime
        } else {
            ExcelDateTimeType::TimeOnly
        };

        self.hour = hour;
        self.min = min;
        self.sec = sec;
        self.datetime_type = date_time_type;

        Ok(self)
    }

    /// Adds to a `ExcelDateTime` date instance with hours, minutes, seconds and
    /// milliseconds.
    ///
    /// Adds time to a existing `ExcelDateTime` date instance or creates a new
    /// one if required.
    ///
    /// # Parameters
    ///
    /// - `hour`: Integer hours in the range 0-23.
    /// - `min`: Integer minutes in the range 0-59.
    /// - `sec`: Integer seconds in the range 0-59.
    /// - `milli`: Integer milliseconds in the range 0-999. Excel only supports
    ///   millisecond precision.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted datetimes in an
    /// Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_and_hms_milli.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    ///     let format4 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.0");
    ///     let format5 = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.000");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms_milli(12, 30, 45, 195)?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_with_format(0, 0, &datetime, &format1)?;
    ///     worksheet.write_with_format(1, 0, &datetime, &format2)?;
    ///     worksheet.write_with_format(2, 0, &datetime, &format3)?;
    ///     worksheet.write_with_format(3, 0, &datetime, &format4)?;
    ///     worksheet.write_with_format(4, 0, &datetime, &format5)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/datetime_and_hms_milli.png">
    ///
    pub fn and_hms_milli(
        mut self,
        hour: u16,
        min: u8,
        sec: u8,
        milli: u16,
    ) -> Result<ExcelDateTime, XlsxError> {
        if let Some(err) = Self::validate_hms_milli(min, sec, milli).err() {
            return Err(err);
        }

        let date_time_type = if self.datetime_type == ExcelDateTimeType::DateOnly {
            ExcelDateTimeType::DateAndTime
        } else {
            ExcelDateTimeType::TimeOnly
        };

        let sec = f64::from(sec) + f64::from(milli) / 1000.0;

        self.hour = hour;
        self.min = min;
        self.sec = sec;
        self.datetime_type = date_time_type;

        Ok(self)
    }

    /// Create a `ExcelDateTime` instance from an Excel serial date.
    ///
    /// An Excel serial date is a f64 number that represents the time since the
    /// Excel epoch. The `from_serial_datetime()` method allows you to create a
    /// `ExcelDateTime` instance from one of these numbers. This is generally
    /// only required if you are creating your own date handling routines or if
    /// you want to manipulate the datetime output from one of the other
    /// routines to account for some offset.
    ///
    /// # Parameters
    ///
    /// - `number`: Excel serial date in the range 0.0 to 2,958,466.0 (years
    ///   1900 to 9999).
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted datetimes in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_from_serial_datetime.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a formats to use with the datetimes below.
    ///     let format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime1 = ExcelDateTime::from_serial_datetime(1.5)?;
    ///     let datetime2 = ExcelDateTime::from_serial_datetime(36526.61)?;
    ///     let datetime3 = ExcelDateTime::from_serial_datetime(44951.72)?;
    ///
    ///     // Write the formatted datetime.
    ///     worksheet.write_with_format(0, 0, &datetime1, &format)?;
    ///     worksheet.write_with_format(1, 0, &datetime2, &format)?;
    ///     worksheet.write_with_format(2, 0, &datetime3, &format)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/datetime_from_serial_datetime.png">
    ///
    pub fn from_serial_datetime(number: impl Into<f64>) -> Result<ExcelDateTime, XlsxError> {
        let number = number.into();
        if !(0.0..2_958_466.0).contains(&number) {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Serial datetime: '{number}' outside converted Excel year range of 1900-9999"
            )));
        }

        let dt = ExcelDateTime {
            serial_datetime: Some(number),
            ..ExcelDateTime::default()
        };

        Ok(dt)
    }

    /// Create a `ExcelDateTime` instance from a Unix time.
    ///
    /// Create a `ExcelDateTime` instance from a [Unix Time] which is the number
    /// of seconds since the 1970-01-01 00:00:00 UTC epoch. This is a common
    /// format used for system times and timestamps.
    ///
    /// Leap seconds are not taken into account.
    ///
    /// [Unix Time]: https://en.wikipedia.org/wiki/Unix_time
    ///
    /// # Parameters
    ///
    /// - `timestamp`: Unix time in the range -2,209,075,200 to 253,402,300,800
    ///   (years 1900 to 9999).
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DateTimeRangeError`] - One of the values used to create the
    ///   date or time is outside Excel's allowed ranges.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted datetimes in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_from_timestamp.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create a formats to use with the datetimes below.
    ///     let format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime1 = ExcelDateTime::from_timestamp(0)?;
    ///     let datetime2 = ExcelDateTime::from_timestamp(1000000000)?;
    ///     let datetime3 = ExcelDateTime::from_timestamp(1687108108)?;
    ///
    ///     // Write the formatted datetime.
    ///     worksheet.write_with_format(0, 0, &datetime1, &format)?;
    ///     worksheet.write_with_format(1, 0, &datetime2, &format)?;
    ///     worksheet.write_with_format(2, 0, &datetime3, &format)?;
    /// #
    /// #     workbook.save("datetime.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/datetime_from_timestamp.png">
    ///
    pub fn from_timestamp(timestamp: i64) -> Result<ExcelDateTime, XlsxError> {
        if !(-2_209_075_200..253_402_300_800).contains(&timestamp) {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Unix timestamp: '{timestamp}' outside converted Excel year range of 1900-9999"
            )));
        }

        // In order to handle negative timestamps in the Excel date range we
        // shift the epoch forward 400 years to get a non-negative timestamp and
        // then subtract 400 years. The 400 years is to get similar leap ranges.
        let timestamp = (UNIX_EPOCH_PLUS_400 + timestamp) as u64;
        let (year, month, day, hour, min, sec) = Self::unix_time_to_date_parts(timestamp);
        let year = year - 400;

        // Validate the date components.
        if let Some(err) = Self::validate_ymd(year, month, day).err() {
            return Err(err);
        }
        if let Some(err) = Self::validate_hms(min, sec).err() {
            return Err(err);
        }

        // Create the new datetime.
        let dt = ExcelDateTime {
            year,
            month,
            day,
            hour,
            min,
            sec,
            datetime_type: ExcelDateTimeType::DateAndTime,
            ..ExcelDateTime::default()
        };

        Ok(dt)
    }

    /// Convert the `ExcelDateTime` to an Excel serial date.
    ///
    /// An Excel serial date is a f64 number that represents the time since the
    /// Excel epoch. This method is mainly used internally when converting an
    /// `ExcelDateTime` instance to an Excel datetime. The method is exposed
    /// publicly to allow some limited manipulation of the date/time in
    /// conjunction with
    /// [`from_serial_datetime()`](ExcelDateTime::from_serial_datetime).
    ///
    /// # Examples
    ///
    /// The following example demonstrates the `ExcelDateTime` `to_excel()`
    /// method.
    ///
    /// ```
    /// # // This code is available in examples/doc_datetime_to_excel.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let time = ExcelDateTime::from_hms(12, 0, 0)?;
    ///     let date = ExcelDateTime::from_ymd(2000, 1, 1)?;
    ///
    ///     assert_eq!(0.5, time.to_excel());
    ///     assert_eq!(36526.0, date.to_excel());
    /// #
    /// #     Ok(())
    /// # }
    ///
    pub fn to_excel(&self) -> f64 {
        if let Some(serial_datetime) = self.serial_datetime {
            serial_datetime
        } else {
            self.to_excel_from_ymd_hms()
        }
    }

    /// Set the Excel date epoch to 1904.
    ///
    /// Excel supports two date epochs: 1900-01-01 and 1904-01-01. The 1904 epoch
    /// has mainly used with Mac for Excel but is a configuration option for
    /// other Excel versions.
    ///
    /// There is some internal support for the 1904 epoch in `ExcelDateTime`
    /// since it was implemented for the Python version of the library. However,
    /// it was almost never used/needed by users so I am omitting it from the
    /// Rust version for now. I won't accept pull requests to implement/unhide
    /// it but I will consider feature requests with a good use
    /// case/justification.
    #[allow(dead_code)]
    pub(crate) fn set_1904_date(mut self) -> ExcelDateTime {
        self.is_1904_date = true;
        self
    }

    // Common validation routine for year, month, day methods.
    fn validate_ymd(year: u16, month: u8, day: u8) -> Result<(), XlsxError> {
        let mut months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        // The default epoch is 1900-01-01 but Excel actually uses 1899-12-31.
        // The upper end of the Excel date range is 9999-12-31.
        if year > 9999 || year < 1900 && (year != 1899 || month != 12 || day != 31) {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Year: '{year}' outside Excel range of 1900-9999"
            )));
        }

        if !(1..=12).contains(&month) {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Month: '{month}' outside Excel range of 1-12"
            )));
        }

        // Change February to account for leap days. Also take Excel's false
        // 1900 leap day into account.
        if Self::is_leap_year(u64::from(year)) || year == 1900 {
            months[1] = 29;
        }

        if day < 1 || day > months[(month as usize) - 1] {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Day: '{day}' is invalid for year '{year}' and month '{month}'"
            )));
        }

        Ok(())
    }

    // Common validation routine for hour, minute, second methods.
    fn validate_hms(min: u8, sec: f64) -> Result<(), XlsxError> {
        // Note, we don't actually validate or restrict the hour. In Excel it
        // can be greater than 24 hours.

        if min > 60 {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Minutes: '{min}' outside Excel range of 0-60"
            )));
        }

        // Excel only supports milli-seconds.
        if sec > 59.999 {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Seconds: '{sec}' outside Excel range of 0-59.999"
            )));
        }

        Ok(())
    }

    // Common validation routine for hour, minute, second, millisecond methods.
    fn validate_hms_milli(min: u8, sec: u8, milli: u16) -> Result<(), XlsxError> {
        // Note, we don't actually validate or restrict the hour. In Excel it
        // can be greater than 24 hours.

        if min > 60 {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Minutes: '{min}' outside Excel range of 0-60"
            )));
        }

        if sec > 60 {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Seconds: '{sec}' outside Excel range of 0-60"
            )));
        }

        // Excel only supports milli-seconds.
        if milli > 999 {
            return Err(XlsxError::DateTimeRangeError(format!(
                "Milliseconds: '{milli}' outside Excel range of 0-999"
            )));
        }

        Ok(())
    }

    // We calculate the date by calculating the number of days since the
    // epoch and adjust for the number of leap days. We calculate the number
    // of leap days by normalizing the year in relation to the epoch. Thus
    // the year 2000 becomes 100 for 4-year and 100-year leapdays and 400
    // for 400-year leapdays.
    pub(crate) fn to_excel_from_ymd_hms(&self) -> f64 {
        let mut year = self.year;
        let mut month = self.month;
        let mut day = self.day;
        let hour = f64::from(self.hour);
        let min = f64::from(self.min);
        let sec = self.sec;
        let is_1904_date = self.is_1904_date;

        let mut months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        let mut days: u32 = 0;
        let mut leap_day = 0;
        let epoch = if is_1904_date { 1904 } else { 1900 };
        let epoch_offset = if is_1904_date { 4 } else { 0 };

        // For times without dates set the default date for the epoch.
        if year == 0 {
            if is_1904_date {
                year = 1904;
                month = 1;
                day = 1;
            } else {
                year = 1899;
                month = 12;
                day = 31;
            }
        }

        // Convert the Excel seconds to a fraction of the seconds in 24 hours.
        let seconds = (hour * 60.0 * 60.0 + min * 60.0 + sec) / (24.0 * 60.0 * 60.0);

        // Special cases for Excel dates in the 1900 epoch.
        if !is_1904_date {
            // The day is on the Excel 1900 epoch (as stored by Excel).
            if year == 1899 && month == 12 && day == 31 {
                return seconds;
            }

            // The day is on the Excel 1900 epoch (another Excel version).
            if year == 1900 && month == 1 && day == 0 {
                return seconds;
            }

            // Excel false leapday. Shortcut the calculations below.
            if year == 1900 && month == 2 && day == 29 {
                return 60.0 + seconds;
            }
        }

        // Normalize the year to the epoch.
        let range = u32::from(year - epoch);

        // Adjust February day count for leap yeas.
        if Self::is_leap_year(u64::from(year)) {
            months[1] = 29;
            leap_day = 1;
        }

        // Add days for previous months.
        for i in 1..month {
            days += months[(i - 1) as usize];
        }

        // Add days for current month.
        days += u32::from(day);

        // Add days for all previous years.
        days += range * 365;

        // Add 4 year leapdays.
        days += range / 4;

        // Remove 100 year leapdays.
        days -= (range + epoch_offset) / 100;

        // Add 400 year leapdays.
        days += (range + epoch_offset + 300) / 400;

        // Remove leap days already counted.
        days -= leap_day;

        // Adjust for Excel erroneously treating 1900 as a leap year.
        if !is_1904_date && days > 59 {
            days += 1;
        }

        f64::from(days) + seconds
    }

    // Convert a Unix time to a ISO 8601 format date.
    //
    // Convert a Unix time (seconds from 1970) to a human readable date in
    // ISO 8601 format.
    pub(crate) fn unix_time_to_rfc3339(timestamp: u64) -> String {
        let (year, month, day, hour, min, sec) = Self::unix_time_to_date_parts(timestamp);

        // Return the ISO 8601 date.
        format!("{year}-{month:02}-{day:02}T{hour:02}:{min:02}:{sec:02}Z",)
    }

    // Convert a Unix time to it date components.
    //
    // The calculation is deceptively tricky since simple division doesn't work
    // due to the 4/100/400 year leap day changes. The basic approach is to
    // divide the range into 400 year blocks, 100 year blocks, 4 year blocks
    // and 1 year block to calculate the year (relative to the epoch). The
    // remaining days and seconds are used to calculate the year day and time.
    //
    // Works in the range 1970-1-1 to 9999-12-31.
    //
    // Leap seconds and the time zone aren't taken into account.
    //
    #[allow(clippy::cast_precision_loss)]
    pub(crate) fn unix_time_to_date_parts(timestamp: u64) -> (u16, u8, u8, u16, u8, f64) {
        let mut months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        // Convert the seconds to a whole number of days.
        let mut days = timestamp / DAY_SECONDS;

        // Move the epoch from 1970 to a 1600 epoch to make the leap
        // calculations easier. This is the closest 400 year epoch before 1970.
        days += 135_140;

        // Get the number of 400 year blocks.
        let year_days_400 = days / YEAR_DAYS_400;
        let mut days = days % YEAR_DAYS_400;

        // Get the number of 100 year blocks. There are 2 kinds: those starting
        // from a %400 year with an extra leap day (36,525 days) and those
        // starting from other 100 year intervals with 1 day less (36,524 days).
        let year_days_100;
        if days < YEAR_DAYS_100 {
            year_days_100 = days / YEAR_DAYS_100;
            days %= YEAR_DAYS_100;
        } else {
            year_days_100 = 1 + (days - YEAR_DAYS_100) / (YEAR_DAYS_100 - 1);
            days = (days - YEAR_DAYS_100) % (YEAR_DAYS_100 - 1);
        }

        // Get the number of 4 year blocks. There are 2 kinds: a 4 year block
        // with a leap day (1461 days) and a 4 year block starting from non-leap
        // %100 years without a leap day (1460 days). We also need to account
        // for whether a 1461 day block was preceded by a 1460 day block at the
        // start of the 100 year block.
        let year_days_4;
        let mut non_leap_year_block = false;
        if year_days_100 == 0 {
            // Any 4 year block in a 36,525 day 100 year block. Has extra leap.
            year_days_4 = days / YEAR_DAYS_4;
            days %= YEAR_DAYS_4;
        } else if days < YEAR_DAYS_4 {
            // A 4 year block at the start of a 36,524 day 100 year block.
            year_days_4 = days / (YEAR_DAYS_4 - 1);
            days %= YEAR_DAYS_4 - 1;
            non_leap_year_block = true;
        } else {
            // A non-initial 4 year block in a 36,524 day 100 year block.
            year_days_4 = 1 + (days - (YEAR_DAYS_4 - 1)) / YEAR_DAYS_4;
            days = (days - (YEAR_DAYS_4 - 1)) % YEAR_DAYS_4;
        }

        // Get the number of 1 year blocks. We need to account for leap years
        // and non-leap years and whether the non-leap occurs after a leap year.
        let year_days_1;
        if non_leap_year_block {
            // A non-leap block not preceded by a leap block.
            year_days_1 = days / YEAR_DAYS;
            days %= YEAR_DAYS;
        } else if days < YEAR_DAYS + 1 {
            // A leap year block.
            year_days_1 = days / (YEAR_DAYS + 1);
            days %= YEAR_DAYS + 1;
        } else {
            // A non-leap block preceded by a leap block.
            year_days_1 = 1 + (days - (YEAR_DAYS + 1)) / YEAR_DAYS;
            days = (days - (YEAR_DAYS + 1)) % YEAR_DAYS;
        }

        // Calculate the year as the number of blocks*days since the epoch.
        let year = 1600 + year_days_400 * 400 + year_days_100 * 100 + year_days_4 * 4 + year_days_1;

        // Convert from 0 indexed to 1 indexed days.
        days += 1;

        // Adjust February day count for leap years.
        if Self::is_leap_year(year) {
            months[1] = 29;
        }

        // Calculate the relevant month.
        let mut month = 1;
        for month_days in months {
            if days > month_days {
                days -= month_days;
                month += 1;
            } else {
                break;
            }
        }

        // The final remainder is the month day.
        let day = days;

        // Get the number of seconds in the day.
        let seconds = timestamp % DAY_SECONDS;

        // Calculate the hours, minutes and seconds in the day.
        let hour = seconds / HOUR_SECONDS;
        let min = (seconds - hour * HOUR_SECONDS) / MINUTE_SECONDS;
        let sec = (seconds - hour * HOUR_SECONDS - min * MINUTE_SECONDS) % MINUTE_SECONDS;

        // Return the date components.
        (
            year as u16,
            month as u8,
            day as u8,
            hour as u16,
            min as u8,
            sec as f64,
        )
    }

    // Check if a year is a leap year.
    pub(crate) fn is_leap_year(year: u64) -> bool {
        year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
    }

    // Get the current UTC time. This is used to set some Excel metadata
    // timestamps.
    pub(crate) fn utc_now() -> String {
        let timestamp = Self::system_now();
        Self::unix_time_to_rfc3339(timestamp)
    }

    // Get the current time from the system time.
    #[cfg(not(all(
        feature = "wasm",
        target_arch = "wasm32",
        not(any(target_os = "emscripten", target_os = "wasi"))
    )))]
    fn system_now() -> u64 {
        let timestamp = SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .expect("SystemTime::now() is before Unix epoch");
        timestamp.as_secs()
    }

    // Get the current time on Wasm/JS systems.
    #[cfg(all(
        feature = "wasm",
        target_arch = "wasm32",
        not(any(target_os = "emscripten", target_os = "wasi"))
    ))]
    fn system_now() -> u64 {
        let timestamp = js_sys::Date::now();
        (timestamp / 1000.0) as u64
    }

    // Convert to UTC date in RFC 3339 format. This is used in custom
    // properties.
    pub(crate) fn to_rfc3339(&self) -> String {
        format!(
            "{}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
            self.year, self.month, self.day, self.hour, self.min, self.sec
        )
    }

    // Chrono date handling functions.

    // Convert a chrono::NaiveTime to an Excel serial datetime.
    #[cfg(feature = "chrono")]
    pub(crate) fn chrono_datetime_to_excel(datetime: &NaiveDateTime) -> f64 {
        let excel_date = Self::chrono_date_to_excel(&datetime.date());
        let excel_time = Self::chrono_time_to_excel(&datetime.time());

        excel_date + excel_time
    }

    // Convert a chrono::NaiveDate to an Excel serial date. In Excel a serial date
    // is the number of days since the epoch, which is either 1899-12-31 or
    // 1904-01-01.
    #[cfg(feature = "chrono")]
    #[allow(clippy::cast_precision_loss)]
    #[allow(clippy::trivially_copy_pass_by_ref)]
    pub(crate) fn chrono_date_to_excel(date: &NaiveDate) -> f64 {
        let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();

        let duration = *date - epoch;
        let mut excel_date = duration.num_days() as f64;

        // For legacy reasons Excel treats 1900 as a leap year. We add an additional
        // day for dates after the leapday in the 1899 epoch.
        if epoch.year() == 1899 && excel_date > 59.0 {
            excel_date += 1.0;
        }

        excel_date
    }

    // Convert a chrono::NaiveTime to an Excel time. The time portion of the Excel
    // datetime is the number of milliseconds divided by the total number of
    // milliseconds in the day.
    #[cfg(feature = "chrono")]
    #[allow(clippy::cast_precision_loss)]
    #[allow(clippy::trivially_copy_pass_by_ref)]
    pub(crate) fn chrono_time_to_excel(time: &NaiveTime) -> f64 {
        let midnight = NaiveTime::from_hms_milli_opt(0, 0, 0, 0).unwrap();
        let duration = *time - midnight;

        duration.num_milliseconds() as f64 / (24.0 * 60.0 * 60.0 * 1000.0)
    }
}

impl Default for ExcelDateTime {
    fn default() -> Self {
        ExcelDateTime {
            year: 0,
            month: 0,
            day: 0,
            hour: 0,
            min: 0,
            sec: 0.0,
            is_1904_date: false,
            serial_datetime: None,
            datetime_type: ExcelDateTimeType::Default,
        }
    }
}

#[derive(Clone, Copy, Eq, PartialEq)]
enum ExcelDateTimeType {
    Default,
    DateOnly,
    TimeOnly,
    DateAndTime,
}

/// Trait to map user date/time types to an Excel serial datetimes.
///
/// The `rust_xlsxwriter` library supports two ways of converting dates and
/// times to Excel dates and times. The first is  via the external [`Chrono`]
/// library which has a comprehensive sets of types and functions for dealing
/// with dates and times. The second is the inbuilt [`ExcelDateTime`] struct
/// which provides a more limited set of methods and which only targets Excel
/// specific dates and times.
///
/// In order to use [`Chrono`] with `rust_xlsxwriter` APIs you must enable the
/// optional `chrono` feature when adding `rust_xlsxwriter` to your
/// `Cargo.toml`.
///
/// [`Chrono`]: https://docs.rs/chrono/latest/chrono
///
pub trait IntoExcelDateTime {
    /// Trait method to convert a date or time into an Excel serial datetime.
    ///
    fn to_excel_serial_date(&self) -> f64;
}

impl IntoExcelDateTime for &ExcelDateTime {
    fn to_excel_serial_date(&self) -> f64 {
        self.to_excel()
    }
}

impl IntoExcelDateTime for ExcelDateTime {
    fn to_excel_serial_date(&self) -> f64 {
        self.to_excel()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelDateTime for &NaiveDateTime {
    fn to_excel_serial_date(&self) -> f64 {
        ExcelDateTime::chrono_datetime_to_excel(self)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelDateTime for &NaiveDate {
    fn to_excel_serial_date(&self) -> f64 {
        ExcelDateTime::chrono_date_to_excel(self)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelDateTime for &NaiveTime {
    fn to_excel_serial_date(&self) -> f64 {
        ExcelDateTime::chrono_time_to_excel(self)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelDateTime for NaiveDateTime {
    fn to_excel_serial_date(&self) -> f64 {
        ExcelDateTime::chrono_datetime_to_excel(self)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelDateTime for NaiveDate {
    fn to_excel_serial_date(&self) -> f64 {
        ExcelDateTime::chrono_date_to_excel(self)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelDateTime for NaiveTime {
    fn to_excel_serial_date(&self) -> f64 {
        ExcelDateTime::chrono_time_to_excel(self)
    }
}

/// Implementation of the `serde::Serialize` trait for `ExcelDateTime`.
///
/// An Excel datetime is a number (see the [`ExcelDateTime`] docs) so it will
/// also need to have an Excel cell format applied to it to display as a date.
///
#[cfg(feature = "serde")]
impl Serialize for ExcelDateTime {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let serial_datetime = self.to_excel();
        serializer.serialize_f64(serial_datetime)
    }
}

/// Implementation of the `serde::Deserialize` trait for `ExcelDateTime`.
///
/// This is a non-functional implementation o allow `ExcelDateTime` types to be
/// included in a struct that derives `Deserialize`.
///
#[cfg(feature = "serde")]
impl<'de> Deserialize<'de> for ExcelDateTime {
    fn deserialize<D>(_deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        unimplemented!()
    }
}
