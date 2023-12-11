// Some utility functions for the `rust_xlsxwriter` module.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Utility functions for `rust_xlsxwriter`.
//!
//! The `rust_xlsxwriter` library provides a number of utility functions for
//! dealing with cell ranges. These can be used for creating formula strings.
//!
//! Note, in general you shouldn't use these functions to create input for APIs
//! that accept string ranges since these is always a similar primary API that
//! accepts numeric ranges.
//!
//! # Examples:
//!
//! ```
//! use rust_xlsxwriter::{cell_range, column_number_to_name};
//!
//! assert_eq!(column_number_to_name(1), "B");
//! assert_eq!(column_number_to_name(702), "AAA");
//!
//! assert_eq!(cell_range(0, 0, 9, 0), "A1:A10");
//! assert_eq!(cell_range(1, 2, 8, 2), "C2:C9");
//! assert_eq!(cell_range(0, 0, 3, 4), "A1:E4");
//! ```

#![warn(missing_docs)]
mod tests;

#[cfg(feature = "serde")]
use crate::IntoExcelDateTime;
#[cfg(feature = "serde")]
use serde::Serializer;

use crate::worksheet::ColNum;
use crate::worksheet::RowNum;
use crate::XlsxError;

/// Convert a zero indexed column cell reference to a string like `"A"`.
///
/// Utility function to convert a zero based column reference to a string
/// representation. This can be useful when constructing ranges for formulas.
///
/// # Examples:
///
/// ```
/// use rust_xlsxwriter::column_number_to_name;
///
/// assert_eq!(column_number_to_name(0), "A");
/// assert_eq!(column_number_to_name(1), "B");
/// assert_eq!(column_number_to_name(702), "AAA");
/// ```
///
pub fn column_number_to_name(col_num: ColNum) -> String {
    let mut col_name = String::new();

    let mut col_num = col_num + 1;

    while col_num > 0 {
        // Set remainder from 1 .. 26
        let mut remainder = col_num % 26;

        if remainder == 0 {
            remainder = 26;
        }

        // Convert the remainder to a character.
        let col_letter = char::from_u32(64u32 + u32::from(remainder)).unwrap();

        // Accumulate the column letters, right to left.
        col_name = format!("{col_letter}{col_name}");

        // Get the next order of magnitude.
        col_num = (col_num - 1) / 26;
    }

    col_name
}

/// Convert a column string such as `"A"` to a zero indexed column reference.
///
/// Utility function to convert a column string representation to a zero based
/// column reference.
///
/// # Examples:
///
/// ```
/// use rust_xlsxwriter::column_name_to_number;
///
/// assert_eq!(column_name_to_number("A"), 0);
/// assert_eq!(column_name_to_number("B"), 1);
/// assert_eq!(column_name_to_number("AAA"), 702);
/// ```
///
pub fn column_name_to_number(column: &str) -> ColNum {
    let mut col_num = 0;

    for char in column.chars() {
        col_num = (col_num * 26) + (char as u16 - 'A' as u16 + 1);
    }

    col_num - 1
}

/// Convert zero indexed row and column cell numbers to a `A1` style string.
///
/// Utility function to convert zero indexed row and column cell values to an
/// `A1` cell reference. This can be useful when constructing ranges for
/// formulas.
///
/// # Examples:
///
/// ```
/// use rust_xlsxwriter::row_col_to_cell;
///
/// assert_eq!(row_col_to_cell(0, 0), "A1");
/// assert_eq!(row_col_to_cell(0, 1), "B1");
/// assert_eq!(row_col_to_cell(1, 1), "B2");
/// ```
///
pub fn row_col_to_cell(row_num: RowNum, col_num: ColNum) -> String {
    format!("{}{}", column_number_to_name(col_num), row_num + 1)
}

/// Convert zero indexed row and column cell numbers to an absolute `$A$1`
/// style range string.
///
/// Utility function to convert zero indexed row and column cell values to an
/// absolute `$A$1` cell reference. This can be useful when constructing ranges
/// for formulas.
///
/// # Examples:
///
/// ```
/// use rust_xlsxwriter::row_col_to_cell_absolute;
///
/// assert_eq!(row_col_to_cell_absolute(0, 0), "$A$1");
/// assert_eq!(row_col_to_cell_absolute(0, 1), "$B$1");
/// assert_eq!(row_col_to_cell_absolute(1, 1), "$B$2");
/// ```
///
pub fn row_col_to_cell_absolute(row_num: RowNum, col_num: ColNum) -> String {
    format!("${}${}", column_number_to_name(col_num), row_num + 1)
}

/// Convert zero indexed row and col cell numbers to a `A1:B1` style range
/// string.
///
/// Utility function to convert zero based row and column cell values to an
/// `A1:B1` style range reference.
///
/// Note, this function should not be used to create a chart range. Use the
/// 5-tuple version of [`IntoChartRange`](crate::IntoChartRange) instead.
///
/// # Examples:
///
/// ```
/// use rust_xlsxwriter::cell_range;
///
/// assert_eq!(cell_range(0, 0, 9, 0), "A1:A10");
/// assert_eq!(cell_range(1, 2, 8, 2), "C2:C9");
/// assert_eq!(cell_range(0, 0, 3, 4), "A1:E4");
/// ```
///
/// If the start and end cell are the same then a single cell range is created:
///
/// ```
/// use rust_xlsxwriter::cell_range;
///
/// assert_eq!(cell_range(0, 0, 0, 0), "A1");
/// ```
///
pub fn cell_range(
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let range1 = row_col_to_cell(first_row, first_col);
    let range2 = row_col_to_cell(last_row, last_col);

    if range1 == range2 {
        range1
    } else {
        format!("{range1}:{range2}")
    }
}

/// Convert zero indexed row and col cell numbers to an absolute `$A$1:$B$1`
/// style range string.
///
/// Utility function to convert zero based row and column cell values to an
/// absolute `$A$1:$B$1` style range reference.
///
/// Note, this function should not be used to create a chart range. Use the
/// 5-tuple version of [`IntoChartRange`](crate::IntoChartRange) instead.
///
/// # Examples:
///
/// ```
/// use rust_xlsxwriter::cell_range_absolute;
///
/// assert_eq!(cell_range_absolute(0, 0, 9, 0), "$A$1:$A$10");
/// assert_eq!(cell_range_absolute(1, 2, 8, 2), "$C$2:$C$9");
/// assert_eq!(cell_range_absolute(0, 0, 3, 4), "$A$1:$E$4");
/// ```
///
/// If the start and end cell are the same then a single cell range is created:
///
/// ```
/// use rust_xlsxwriter::cell_range_absolute;
///
/// assert_eq!(cell_range_absolute(0, 0, 0, 0), "$A$1");
/// ```
///
pub fn cell_range_absolute(
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let range1 = row_col_to_cell_absolute(first_row, first_col);
    let range2 = row_col_to_cell_absolute(last_row, last_col);

    if range1 == range2 {
        range1
    } else {
        format!("{range1}:{range2}")
    }
}

/// Serialize a `Chrono` naive date/time to and Excel value.
///
/// This is a helper function for serializing [`Chrono`] naive (i.e., timezone
/// unaware) date/time fields using [Serde](https://serde.rs). See [Working with
/// Serde](crate::serializer#working-with-serde) for more information.
///
/// The function works for the following types:
///   - [`NaiveDateTime`]
///   - [`NaiveDate`]
///   - [`NaiveTime`]
///
/// [`Chrono`]: https://docs.rs/chrono/latest/chrono
/// [`NaiveDate`]:
///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
/// [`NaiveTime`]:
///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
/// [`NaiveDateTime`]:
///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
///
/// # Errors
///
/// * [`XlsxError::SerdeError`] - A wrapped serialization error.
///
#[cfg(feature = "serde")]
pub fn serialize_chrono_naive_to_excel<S>(
    datetime: impl IntoExcelDateTime,
    serializer: S,
) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    serializer.serialize_f64(datetime.to_excel_serial_date())
}

// Convert zero indexed row and col cell references to a chart absolute
// Sheet1!$A$1:$B$1 style range string.
pub(crate) fn chart_range_abs(
    sheet_name: &str,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let sheet_name = quote_sheetname(sheet_name);
    let range1 = row_col_to_cell_absolute(first_row, first_col);
    let range2 = row_col_to_cell_absolute(last_row, last_col);

    if range1 == range2 {
        format!("{sheet_name}!{range1}")
    } else {
        format!("{sheet_name}!{range1}:{range2}")
    }
}

// Create a quoted version of a worksheet name. Excel single quotes worksheet
// names that contain spaces and some other characters.
pub(crate) fn quote_sheetname(sheetname: &str) -> String {
    let mut sheetname = sheetname.to_string();

    // Ignore strings that are already quoted.
    if !sheetname.starts_with('\'') {
        // double quote and other single quotes.
        sheetname = sheetname.replace('\'', "''");

        // Single quote the worksheet name if it contains any of the characters
        // that Excel quotes when using the name in a formula.
        if sheetname.contains(' ') || sheetname.contains('!') || sheetname.contains('\'') {
            sheetname = format!("'{sheetname}'");
        }
    }

    sheetname
}

pub(crate) fn validate_sheetname(name: &str, message: &str) -> Result<(), XlsxError> {
    // Check that the sheet name isn't blank.
    if name.is_empty() {
        return Err(XlsxError::SheetnameCannotBeBlank(message.to_string()));
    }

    // Check that sheet sheetname is <= 31, an Excel limit.
    if name.chars().count() > 31 {
        return Err(XlsxError::SheetnameLengthExceeded(message.to_string()));
    }

    // Check that sheetname doesn't contain any invalid characters.
    if name.contains(['*', '?', ':', '[', ']', '\\', '/']) {
        return Err(XlsxError::SheetnameContainsInvalidCharacter(
            message.to_string(),
        ));
    }

    // Check that sheetname doesn't start or end with an apostrophe.
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(XlsxError::SheetnameStartsOrEndsWithApostrophe(
            message.to_string(),
        ));
    }

    Ok(())
}

// Get the pixel width of a string based on character widths taken from Excel.
// Non-ascii characters are given a default width of 8 pixels.
#[allow(clippy::match_same_arms)]
pub(crate) fn pixel_width(string: &str) -> u16 {
    let mut length = 0;

    for char in string.chars() {
        match char {
            ' ' | '\'' => length += 3,

            ',' | '.' | ':' | ';' | 'I' | '`' | 'i' | 'j' | 'l' => length += 4,

            '!' | '(' | ')' | '-' | 'J' | '[' | ']' | 'f' | 'r' | 't' | '{' | '}' => length += 5,

            '"' | '/' | 'L' | '\\' | 'c' | 's' | 'z' => length += 6,

            '#' | '$' | '*' | '+' | '0' | '1' | '2' | '3' | '4' | '5' | '6' | '7' | '8' | '9'
            | '<' | '=' | '>' | '?' | 'E' | 'F' | 'S' | 'T' | 'Y' | 'Z' | '^' | '_' | 'a' | 'g'
            | 'k' | 'v' | 'x' | 'y' | '|' | '~' => length += 7,

            'B' | 'C' | 'K' | 'P' | 'R' | 'X' | 'b' | 'd' | 'e' | 'h' | 'n' | 'o' | 'p' | 'q'
            | 'u' => length += 8,

            'A' | 'D' | 'G' | 'H' | 'U' | 'V' => length += 9,

            '&' | 'N' | 'O' | 'Q' => length += 10,

            '%' | 'w' => length += 11,

            'M' | 'm' => length += 12,

            '@' | 'W' => length += 13,

            _ => length += 8,
        }
    }

    length
}

// Hash a worksheet password. Based on the algorithm in ECMA-376-4:2016, Office
// Open XML File Formats — Transitional Migration Features, Additional
// attributes for workbookProtection element (Part 1, §18.2.29).
pub(crate) fn hash_password(password: &str) -> u16 {
    let mut hash: u16 = 0;
    let length = password.len() as u16;

    if password.is_empty() {
        return 0;
    }

    for byte in password.as_bytes().iter().rev() {
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
        hash ^= u16::from(*byte);
    }

    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= length;
    hash ^= 0xCE4B;

    hash
}

// Clone and strip the leading '=' from formulas, if present.
pub(crate) fn formula_to_string(formula: &str) -> String {
    let mut formula = formula.to_string();

    if formula.starts_with('=') {
        formula.remove(0);
    }

    formula
}

// Trait to convert bool to XML "0" or "1".
pub(crate) trait ToXmlBoolean {
    fn to_xml_bool(self) -> String;
}

impl ToXmlBoolean for bool {
    fn to_xml_bool(self) -> String {
        u8::from(self).to_string()
    }
}
