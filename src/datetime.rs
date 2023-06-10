// excel_datetime - A module for handling Excel dates and times.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
use std::time::SystemTime;

/// A struct to represent an Excel date and/or time.
///
/// TODO.
pub struct ExcelDateTime {}

const DAY_SECONDS: u64 = 24 * 60 * 60;
const HOUR_SECONDS: u64 = 60 * 60;
const MINUTE_SECONDS: u64 = 60;
const YEAR_DAYS: u64 = 365;
const YEAR_DAYS_4: u64 = YEAR_DAYS * 4 + 1;
const YEAR_DAYS_100: u64 = YEAR_DAYS * 100 + 25;
const YEAR_DAYS_400: u64 = YEAR_DAYS * 400 + 97;

impl ExcelDateTime {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new `ExcelDateTime` struct instance.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ExcelDateTime {
        ExcelDateTime {}
    }

    /// Convert a Unix time to a ISO8601 format date.
    ///
    /// Convert a Unix time (seconds from 1970) to a human readable date in
    /// ISO8601 format.
    ///
    /// The calculation is deceptively tricky since simple division doesn't work
    /// due to the 4/100/400 year leap day changes. The basic approach is to
    /// divide the range into 400 year blocks, 100 year blocks, 4 year blocks
    /// and 1 year block to calculate the year (relative to the epoch). The
    /// remaining days and seconds are used to calculate the year day and time.
    ///
    /// Works in the range 1970-1-1 to 9999-12-31.
    ///
    /// Leap seconds and the time zone aren't taken into account.
    ///
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
    #[allow(clippy::match_wild_err_arm)] // todo
    pub(crate) fn utc_now() -> u64 {
        match SystemTime::now().duration_since(SystemTime::UNIX_EPOCH) {
            Ok(n) => n.as_secs(),
            // TODO add better error handling here.
            Err(_) => panic!("SystemTime before UNIX EPOCH!"),
        }
    }
}

// -----------------------------------------------------------------------
// Tests are in the datetime sub-directory.
// -----------------------------------------------------------------------
mod tests;
