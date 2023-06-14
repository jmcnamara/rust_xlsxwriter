// excel_datetime - A module for handling Excel dates and times.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
use regex::Regex;
use std::time::SystemTime;

use crate::XlsxError;

const DAY_SECONDS: u64 = 24 * 60 * 60;
const HOUR_SECONDS: u64 = 60 * 60;
const MINUTE_SECONDS: u64 = 60;
const YEAR_DAYS: u64 = 365;
const YEAR_DAYS_4: u64 = YEAR_DAYS * 4 + 1;
const YEAR_DAYS_100: u64 = YEAR_DAYS * 100 + 25;
const YEAR_DAYS_400: u64 = YEAR_DAYS * 400 + 97;

/// A struct to represent an Excel date and/or time.
///
/// TODO.
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
    num_format: String,
}

impl ExcelDateTime {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
    ///
    pub fn parse_from_str(datetime: &str) -> Result<ExcelDateTime, XlsxError> {
        lazy_static! {
            static ref DATE: Regex = Regex::new(r"\b(\d\d\d\d)-(\d\d)-(\d\d)").unwrap();
            static ref TIME: Regex = Regex::new(r"(\d+):(\d\d)(:(\d\d(\.\d+)?))?").unwrap();
        }
        let mut matched = false;

        let mut dt = match DATE.captures(datetime) {
            Some(caps) => {
                let year = caps.get(1).unwrap().as_str().parse::<u16>().unwrap();
                let month = caps.get(2).unwrap().as_str().parse::<u8>().unwrap();
                let day = caps.get(3).unwrap().as_str().parse::<u8>().unwrap();

                matched = true;
                ExcelDateTime::from_ymd(year, month, day)
            }
            None => Ok(ExcelDateTime::default()),
        };

        if let Some(caps) = TIME.captures(datetime) {
            let hour = caps.get(1).unwrap().as_str().parse::<u16>().unwrap();
            let min = caps.get(2).unwrap().as_str().parse::<u8>().unwrap();

            let sec = match caps.get(3) {
                Some(_) => caps.get(4).unwrap().as_str().parse::<f64>().unwrap(),
                None => 0.0,
            };

            matched = true;
            dt = dt.unwrap().and_hms(hour, min, sec);
        }

        if !matched {
            return Err(XlsxError::DateParseError(datetime.to_string()));
        }

        dt
    }

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
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

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
    ///
    pub fn from_hms(hour: u16, min: u8, sec: impl Into<f64>) -> Result<ExcelDateTime, XlsxError> {
        ExcelDateTime::default().and_hms(hour, min, sec)
    }

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
    ///
    pub fn from_hms_milli(
        hour: u16,
        min: u8,
        sec: u8,
        milli: u16,
    ) -> Result<ExcelDateTime, XlsxError> {
        ExcelDateTime::default().and_hms_milli(hour, min, sec, milli)
    }

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
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

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
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

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
    ///
    pub fn from_serial_datetime(number: impl Into<f64>) -> Result<ExcelDateTime, XlsxError> {
        let number = number.into();
        if !(0.0..2_958_466.0).contains(&number) {
            return Err(XlsxError::DateRangeError(format!(
                "Serial datetime: '{number}' outside converted Excel year range of 1900-9999"
            )));
        }

        let dt = ExcelDateTime {
            serial_datetime: Some(number),
            ..ExcelDateTime::default()
        };

        Ok(dt)
    }

    /// Create a `ExcelDateTime` instance from TODO.
    ///
    /// # Errors
    ///
    /// TODO
    ///
    #[allow(clippy::cast_precision_loss)]
    pub fn from_timestamp(timestamp: i64) -> Result<ExcelDateTime, XlsxError> {
        if !(-2_209_075_200..253_402_300_800).contains(&timestamp) {
            return Err(XlsxError::DateRangeError(format!(
                "Unix timestamp: '{timestamp}' outside converted Excel year range of 1900-9999"
            )));
        }

        let days = (timestamp / (24 * 60 * 60)) as f64;
        let time = ((timestamp % (24 * 60 * 60)) as f64) / (24.0 * 60.0 * 60.0);
        let mut datetime = 25568.0 + days + time;

        if datetime >= 60.0 {
            datetime += 1.0;
        }

        let dt = ExcelDateTime {
            serial_datetime: Some(datetime),
            ..ExcelDateTime::default()
        };

        Ok(dt)
    }

    /// TODO
    pub fn set_num_format(mut self, num_format: impl Into<String>) -> ExcelDateTime {
        self.num_format = num_format.into();
        self
    }

    // TODO
    #[allow(dead_code)]
    pub(crate) fn set_1904_date(mut self) -> ExcelDateTime {
        self.is_1904_date = true;
        self
    }

    // TODO
    pub(crate) fn get_num_format(&self) -> String {
        if self.num_format.is_empty() {
            match self.datetime_type {
                ExcelDateTimeType::DateOnly => String::from("yyyy\\-mm\\-dd;@"),
                ExcelDateTimeType::TimeOnly => String::from("hh:mm:ss;@"),
                ExcelDateTimeType::DateAndTime | ExcelDateTimeType::Default => {
                    String::from("yyyy\\-mm\\-dd\\ hh:mm:ss")
                }
            }
        } else {
            self.num_format.clone()
        }
    }

    // TODO
    fn validate_ymd(year: u16, month: u8, day: u8) -> Result<(), XlsxError> {
        let mut months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        // The default epoch is 1900-01-01 but Excel actually uses 1899-12-31.
        // The upper end of the Excel date range is 9999-12-31.
        if year > 9999 || year < 1900 && (year != 1899 || month != 12 || day != 31) {
            return Err(XlsxError::DateRangeError(format!(
                "Year: '{year}' outside Excel range of 1900-9999"
            )));
        }

        if !(1..=12).contains(&month) {
            return Err(XlsxError::DateRangeError(format!(
                "Month: '{month}' outside Excel range of 1-12"
            )));
        }

        // Change February to account for leap days. Also take Excel's false
        // 1900 leap day into account.
        if Self::is_leap_year(u64::from(year)) || year == 1900 {
            months[1] = 29;
        }

        if day < 1 || day > months[(month as usize) - 1] {
            return Err(XlsxError::DateRangeError(format!(
                "Day: '{day}' is invalid for year '{year}' and month '{month}'"
            )));
        }

        Ok(())
    }

    // TODO
    fn validate_hms(min: u8, sec: f64) -> Result<(), XlsxError> {
        // Note, we don't actually validate or restrict the hour. In Excel it
        // can be greater than 24 hours.

        if min > 60 {
            return Err(XlsxError::DateRangeError(format!(
                "Minutes: '{min}' outside Excel range of 0-60"
            )));
        }

        // Excel only supports milli-seconds.
        if sec > 59.999 {
            return Err(XlsxError::DateRangeError(format!(
                "Seconds: '{sec}' outside Excel range of 0-59.999"
            )));
        }

        Ok(())
    }

    // TODO
    fn validate_hms_milli(min: u8, sec: u8, milli: u16) -> Result<(), XlsxError> {
        // Note, we don't actually validate or restrict the hour. In Excel it
        // can be greater than 24 hours.

        if min > 60 {
            return Err(XlsxError::DateRangeError(format!(
                "Minutes: '{min}' outside Excel range of 0-60"
            )));
        }

        if sec > 60 {
            return Err(XlsxError::DateRangeError(format!(
                "Seconds: '{sec}' outside Excel range of 0-60"
            )));
        }

        // Excel only supports milli-seconds.
        if milli > 999 {
            return Err(XlsxError::DateRangeError(format!(
                "Milliseconds: '{milli}' outside Excel range of 0-999"
            )));
        }

        Ok(())
    }

    /// TODO.
    pub fn to_excel(&self) -> f64 {
        if let Some(serial_datetime) = self.serial_datetime {
            serial_datetime
        } else {
            self.to_excel_from_ymd_hms()
        }
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

    // Convert a Unix time to a ISO8601 format date.
    //
    // Convert a Unix time (seconds from 1970) to a human readable date in
    // ISO8601 format.
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
    pub(crate) fn unix_time_to_iso8601(timestamp: u64) -> String {
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

        // Return the ISO8601 date.
        format!("{year}-{month:02}-{day:02}T{hour:02}:{min:02}:{sec:02}Z",)
    }

    // Check if a year is a leap year.
    pub(crate) fn is_leap_year(year: u64) -> bool {
        year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
    }

    // TODO
    pub(crate) fn utc_now() -> u64 {
        let timestamp = SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .expect("SystemTime::now() is before Unix epoch");

        timestamp.as_secs()
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
            num_format: String::new(),
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

// -----------------------------------------------------------------------
// Tests are in the datetime sub-directory.
// -----------------------------------------------------------------------
mod tests;
