// Some utility functions for the `rust_xlsxwriter` module.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Utility functions for `rust_xlsxwriter`.
//!
//! The `rust_xlsxwriter` library provides a number of utility functions for
//! dealing with cell ranges, Chrono Serde serialization, and other helper
//! method.
//!
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
use crate::COL_MAX;
use crate::ROW_MAX;
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
    if column.is_empty() {
        return 0;
    }

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

/// Serialize a Chrono naive date/time to an Excel value.
///
/// This is a helper function for serializing [`Chrono`] naive date/time fields
/// using [Serde](https://serde.rs). "Naive" in the Chrono sense means that the
/// dates/times don't have timezone information, like Excel.
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
/// `Option<T>` Chrono types can be handled with
/// [`serialize_chrono_option_naive_to_excel()`].
///
/// See [Working with Serde](crate::serializer#working-with-serde) for more
/// information about serialization with `rust_xlsxwriter`.
///
/// # Errors
///
/// - [`XlsxError::SerdeError`] - A wrapped serialization error.
///
/// # Examples
///
/// Example of a serializable struct with a Chrono Naive value with a helper
/// function.
///
/// ```
/// # // This code is available in examples/doc_worksheet_serialize_datetime3.rs
/// #
/// use rust_xlsxwriter::utility::serialize_chrono_naive_to_excel;
/// use serde::Serialize;
///
/// fn main() {
///     #[derive(Serialize)]
///     struct Student {
///         full_name: String,
///
///         #[serde(serialize_with = "serialize_chrono_naive_to_excel")]
///         birth_date: NaiveDate,
///
///         id_number: u32,
///     }
/// }
/// ```
///
#[cfg(feature = "serde")]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub fn serialize_chrono_naive_to_excel<S>(
    datetime: impl IntoExcelDateTime,
    serializer: S,
) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    serializer.serialize_f64(datetime.to_excel_serial_date())
}

/// Serialize an `Option` Chrono naive date/time to an Excel value.
///
/// This is a helper function for serializing [`Chrono`] naive date/time fields
/// using [Serde](https://serde.rs). "Naive" in the Chrono sense means that the
/// dates/times don't have timezone information, like Excel.
///
/// A helper function is provided for [`Option`] Chrono values since it is
/// common to have `Option<NaiveDate>` values as a result of deserialization. It
/// also takes care of the use case where you want a `None` value to be written
/// as a blank cell with the same cell format as other values of the field type.
///
/// The function works for the following `Option<T>` where T is:
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
/// Non `Option<T>` Chrono types can be handled with
/// [`serialize_chrono_naive_to_excel()`].
///
/// See [Working with Serde](crate::serializer#working-with-serde) for more
/// information about serialization with `rust_xlsxwriter`.
///
/// # Errors
///
/// - [`XlsxError::SerdeError`] - A wrapped serialization error.
///
/// # Examples
///
/// Example of a serializable struct with an Option Chrono Naive value with a
/// helper function.
///
///
/// ```
/// # // This code is available in examples/doc_worksheet_serialize_datetime5.rs
/// #
/// use rust_xlsxwriter::utility::serialize_chrono_option_naive_to_excel;
/// use serde::Serialize;
///
/// fn main() {
///     #[derive(Serialize)]
///     struct Student {
///         full_name: String,
///
///         #[serde(serialize_with = "serialize_chrono_option_naive_to_excel")]
///         birth_date: Option<NaiveDate>,
///
///         id_number: u32,
///     }
/// }
/// ```
///
#[cfg(feature = "serde")]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub fn serialize_chrono_option_naive_to_excel<S>(
    datetime: &Option<impl IntoExcelDateTime>,
    serializer: S,
) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    match datetime {
        Some(datetime) => serializer.serialize_f64(datetime.to_excel_serial_date()),
        None => serializer.serialize_none(),
    }
}

// Convert zero indexed row and col cell references to a non-absolute chart
// "Sheet1!A1:B1" style range string.
pub(crate) fn chart_range(
    sheet_name: &str,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let sheet_name = quote_sheetname(sheet_name);
    let range1 = row_col_to_cell(first_row, first_col);
    let range2 = row_col_to_cell(last_row, last_col);

    if range1 == range2 {
        format!("{sheet_name}!{range1}")
    } else {
        format!("{sheet_name}!{range1}:{range2}")
    }
}

// Convert zero indexed row and col cell references to an absolute chart
// "Sheet1!$A$1:$B$1" style range string.
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

// Convert zero indexed row and col cell references to a range and tuple string
// suitable for an error message.
pub(crate) fn chart_error_range(
    sheet_name: &str,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let sheet_name = quote_sheetname(sheet_name);
    let range1 = row_col_to_cell(first_row, first_col);
    let range2 = row_col_to_cell(last_row, last_col);

    if range1 == range2 {
        format!("{sheet_name}!{range1}/({first_row}, {first_col})")
    } else {
        format!("{sheet_name}!{range1}:{range2}/({first_row}, {first_col}, {last_row}, {last_col})")
    }
}

// Sheetnames used in references should be quoted if they contain any spaces,
// special characters or if they look like a A1 or RC cell reference. The rules
// are shown inline below.
#[allow(clippy::if_same_then_else)]
pub(crate) fn quote_sheetname(sheetname: &str) -> String {
    let mut sheetname = sheetname.to_string();
    let uppercase_sheetname = sheetname.to_uppercase();
    let mut requires_quoting = false;
    let col_max = u64::from(COL_MAX);
    let row_max = u64::from(ROW_MAX);

    // Split sheetnames that look like A1 and R1C1 style cell references into a
    // leading string and a trailing number.
    let (string_part, number_part) = split_cell_reference(&sheetname);

    // The number part of the sheet name can have trailing non-digit characters
    // and still be a valid R1C1 match. However, to test the R1C1 row/col part
    // we need to extract just the number part.
    let mut number_parts = number_part.split(|c: char| !c.is_ascii_digit());
    let rc_number_part = number_parts.next().unwrap_or_default();

    // Ignore strings that are already quoted.
    if !sheetname.starts_with('\'') {
        // --------------------------------------------------------------------
        // Rule 1. Sheet names that contain anything other than \w and "."
        // characters must be quoted.
        // --------------------------------------------------------------------

        if !sheetname
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.' || is_emoji(c))
        {
            requires_quoting = true;
        }
        // --------------------------------------------------------------------
        // Rule 2. Sheet names that start with a digit or "." must be quoted.
        // --------------------------------------------------------------------
        else if sheetname.starts_with(|c: char| c.is_ascii_digit() || c == '.' || is_emoji(c)) {
            requires_quoting = true;
        }
        // --------------------------------------------------------------------
        // Rule 3. Sheet names must not be a valid A1 style cell reference.
        // Valid means that the row and column range values must also be within
        // Excel row and column limits.
        // --------------------------------------------------------------------
        else if (1..=3).contains(&string_part.len())
            && number_part.chars().all(|c| c.is_ascii_digit())
        {
            let col = column_name_to_number(&string_part);
            let col = u64::from(col + 1);

            let row = number_part.parse::<u64>().unwrap_or_default();

            if row > 0 && row <= row_max && col <= col_max {
                requires_quoting = true;
            }
        }
        // --------------------------------------------------------------------
        // Rule 4. Sheet names must not *start* with a valid RC style cell
        // reference. Other characters after the valid RC reference are ignored
        // by Excel. Valid means that the row and column range values must also
        // be within Excel row and column limits.
        //
        // Note: references without trailing characters like R12345 or C12345
        // are caught by Rule 3. Negative references like R-12345 are caught by
        // Rule 1 due to dash.
        // --------------------------------------------------------------------

        // Rule 4a. Check for sheet names that start with R1 style references.
        else if string_part == "R" {
            let row = rc_number_part.parse::<u64>().unwrap_or_default();

            if row > 0 && row <= row_max {
                requires_quoting = true;
            }
        }
        // Rule 4b. Check for sheet names that start with C1 or RC1 style
        // references.
        else if string_part == "RC" || string_part == "C" {
            let col = rc_number_part.parse::<u64>().unwrap_or_default();

            if col > 0 && col <= col_max {
                requires_quoting = true;
            }
        }
        // Rule 4c. Check for some single R/C references.
        else if uppercase_sheetname == "R"
            || uppercase_sheetname == "C"
            || uppercase_sheetname == "RC"
        {
            requires_quoting = true;
        }
    }

    if requires_quoting {
        // Double up any single quotes.
        sheetname = sheetname.replace('\'', "''");

        // Single quote the sheet name.
        sheetname = format!("'{sheetname}'");
    }

    sheetname
}

// Unquote an Excel single quoted string.
pub(crate) fn unquote_sheetname(sheetname: &str) -> String {
    if sheetname.starts_with('\'') && sheetname.ends_with('\'') {
        let sheetname = sheetname[1..sheetname.len() - 1].to_string();
        sheetname.replace("''", "'")
    } else {
        sheetname.to_string()
    }
}

// Match emoji characters when quoting sheetnames. The following were generated from:
// https://util.unicode.org/UnicodeJsps/list-unicodeset.jsp?a=%5B%3AEmoji%3DYes%3A%5D&abb=on&esc=on&g=&i=
//
pub(crate) fn is_emoji(c: char) -> bool {
    if c < '\u{203C}' {
        // Shortcut for most chars in the lower range. We ignore chars '#', '*',
        // '0-9', '©️' and '®️' which are in the range and which are, strictly
        // speaking, emoji symbols, but they are not treated so by Excel in the
        // context of this check.
        return false;
    }

    if c < '\u{01F004}' {
        return matches!(c,
            '\u{203C}' | '\u{2049}' | '\u{2122}' | '\u{2139}' | '\u{2194}'..='\u{2199}' |
            '\u{21A9}' | '\u{21AA}' | '\u{231A}' | '\u{231B}' | '\u{2328}' | '\u{23CF}' |
            '\u{23E9}'..='\u{23F3}' | '\u{23F8}'..='\u{23FA}' | '\u{24C2}' | '\u{25AA}' |
            '\u{25AB}' | '\u{25B6}' | '\u{25C0}' | '\u{25FB}'..='\u{25FE}' |
            '\u{2600}'..='\u{2604}' | '\u{260E}' | '\u{2611}' | '\u{2614}' | '\u{2615}' |
            '\u{2618}' | '\u{261D}' | '\u{2620}' | '\u{2622}' | '\u{2623}' | '\u{2626}' |
            '\u{262A}' | '\u{262E}' | '\u{262F}' | '\u{2638}'..='\u{263A}' | '\u{2640}' |
            '\u{2642}' | '\u{2648}'..='\u{2653}' | '\u{265F}' | '\u{2660}' | '\u{2663}' |
            '\u{2665}' | '\u{2666}' | '\u{2668}' | '\u{267B}' | '\u{267E}' | '\u{267F}' |
            '\u{2692}'..='\u{2697}' | '\u{2699}' | '\u{269B}' | '\u{269C}' | '\u{26A0}' |
            '\u{26A1}' | '\u{26A7}' | '\u{26AA}' | '\u{26AB}' | '\u{26B0}' | '\u{26B1}' |
            '\u{26BD}' | '\u{26BE}' | '\u{26C4}' | '\u{26C5}' | '\u{26C8}' | '\u{26CE}' |
            '\u{26CF}' | '\u{26D1}' | '\u{26D3}' | '\u{26D4}' | '\u{26E9}' | '\u{26EA}' |
            '\u{26F0}'..='\u{26F5}' | '\u{26F7}'..='\u{26FA}' | '\u{26FD}' | '\u{2702}' |
            '\u{2705}' | '\u{2708}'..='\u{270D}' | '\u{270F}' | '\u{2712}' | '\u{2714}' |
            '\u{2716}' | '\u{271D}' | '\u{2721}' | '\u{2728}' | '\u{2733}' | '\u{2734}' |
            '\u{2744}' | '\u{2747}' | '\u{274C}' | '\u{274E}' | '\u{2753}'..='\u{2755}' |
            '\u{2757}' | '\u{2763}' | '\u{2764}' | '\u{2795}'..='\u{2797}' | '\u{27A1}' |
            '\u{27B0}' | '\u{27BF}' | '\u{2934}' | '\u{2935}' | '\u{2B05}'..='\u{2B07}' |
            '\u{2B1B}' | '\u{2B1C}' | '\u{2B50}' | '\u{2B55}' | '\u{3030}' | '\u{303D}' |
            '\u{3297}' | '\u{3299}'
        );
    }

    matches!(c,
        '\u{01F004}' | '\u{01F0CF}' | '\u{01F170}' | '\u{01F171}' | '\u{01F17E}' | '\u{01F17F}' |
        '\u{01F18E}' | '\u{01F191}'..='\u{01F19A}' | '\u{01F1E6}'..='\u{01F1FF}' | '\u{01F201}' |
        '\u{01F202}' | '\u{01F21A}' | '\u{01F22F}' | '\u{01F232}'..='\u{01F23A}' | '\u{01F250}' |
        '\u{01F251}' | '\u{01F300}'..='\u{01F321}' | '\u{01F324}'..='\u{01F393}' | '\u{01F396}' |
        '\u{01F397}' | '\u{01F399}'..='\u{01F39B}' | '\u{01F39E}'..='\u{01F3F0}' |
        '\u{01F3F3}'..='\u{01F3F5}' | '\u{01F3F7}'..='\u{01F4FD}' | '\u{01F4FF}'..='\u{01F53D}' |
        '\u{01F549}'..='\u{01F54E}' | '\u{01F550}'..='\u{01F567}' | '\u{01F56F}' | '\u{01F570}' |
        '\u{01F573}'..='\u{01F57A}' | '\u{01F587}' | '\u{01F58A}'..='\u{01F58D}' | '\u{01F590}' |
        '\u{01F595}' | '\u{01F596}' | '\u{01F5A4}' | '\u{01F5A5}' | '\u{01F5A8}' | '\u{01F5B1}' |
        '\u{01F5B2}' | '\u{01F5BC}' | '\u{01F5C2}'..='\u{01F5C4}' | '\u{01F5D1}'..='\u{01F5D3}' |
        '\u{01F5DC}'..='\u{01F5DE}' | '\u{01F5E1}' | '\u{01F5E3}' | '\u{01F5E8}' | '\u{01F5EF}' |
        '\u{01F5F3}' | '\u{01F5FA}'..='\u{01F64F}' | '\u{01F680}'..='\u{01F6C5}' |
        '\u{01F6CB}'..='\u{01F6D2}' | '\u{01F6D5}'..='\u{01F6D7}' | '\u{01F6DC}'..='\u{01F6E5}' |
        '\u{01F6E9}' | '\u{01F6EB}' | '\u{01F6EC}' | '\u{01F6F0}' | '\u{01F6F3}'..='\u{01F6FC}' |
        '\u{01F7E0}'..='\u{01F7EB}' | '\u{01F7F0}' | '\u{01F90C}'..='\u{01F93A}' |
        '\u{01F93C}'..='\u{01F945}' | '\u{01F947}'..='\u{01F9FF}' | '\u{01FA70}'..='\u{01FA7C}' |
        '\u{01FA80}'..='\u{01FA88}' | '\u{01FA90}'..='\u{01FABD}' | '\u{01FABF}'..='\u{01FAC5}' |
        '\u{01FACE}'..='\u{01FADB}' | '\u{01FAE0}'..='\u{01FAE8}' | '\u{01FAF0}'..='\u{01FAF8}'
    )
}

// Split sheetnames that look like A1 and R1C1 style cell references into a
// leading string and a trailing number.
pub(crate) fn split_cell_reference(sheetname: &str) -> (String, String) {
    match sheetname.find(|c: char| c.is_ascii_digit()) {
        Some(position) => (
            (sheetname[..position]).to_uppercase(),
            (sheetname[position..]).to_uppercase(),
        ),
        None => (String::new(), String::new()),
    }
}

// Check that a range string like "A1" or "A1:B3" are valid. This function
// assumes that the '$' absolute anchor has already been stripped.
pub(crate) fn is_valid_range(range: &str) -> bool {
    if range.is_empty() {
        return false;
    }

    // The range should start with a letter and end in a number.
    if !range.starts_with(|c: char| c.is_ascii_uppercase())
        || !range.ends_with(|c: char| c.is_ascii_digit())
    {
        return false;
    }

    // The range should only include the characters 'A-Z', '0-9' and ':'
    if !range
        .chars()
        .all(|c: char| c.is_ascii_uppercase() || c.is_ascii_digit() || c == ':')
    {
        return false;
    }

    true
}

/// Check that a worksheet name is valid in Excel.
///
/// This function checks if an worksheet name is valid according to the Excel
/// rules:
///
/// - The name is less than 32 characters.
/// - The name isn't blank.
/// - The name doesn't contain any of the characters: `[ ] : * ? / \`.
/// - The name doesn't start or end with an apostrophe.
///
/// The worksheet name "History" isn't allowed in English versions of Excel
/// since it is a reserved name. However it is allowed in some other language
/// versions so this function doesn't raise it as an error. Overall it is best
/// to avoid using it.
///
/// The rules for worksheet names in Excel are explained in the [Microsoft
/// Office documentation].
///
/// [Microsoft Office documentation]:
///     https://support.office.com/en-ie/article/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
///
/// # Parameters
///
/// - `name`: The worksheet name. It must follow the Excel rules, shown above.
///
/// # Errors
///
/// - [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name cannot be blank.
/// - [`XlsxError::SheetnameLengthExceeded`] - Worksheet name exceeds Excel's
///   limit of 31 characters.
/// - [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name cannot
///   contain invalid characters: `[ ] : * ? / \`
/// - [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name cannot
///   start or end with an apostrophe.
///
/// # Examples
///
/// The following example demonstrates testing for a valid worksheet name.
///
/// ```fail
/// # // This code is available in examples/doc_utility_check_sheet_name.rs
/// #
/// # use rust_xlsxwriter::{utility, XlsxError};
/// #
/// fn main() -> Result<(), XlsxError> {
///     // This worksheet name is valid.
///     utility::check_sheet_name("2030-01-01")?;
///
///     // This worksheet name isn't valid due to the forward slashes.
///     utility::check_sheet_name("2030/01/01")?;
///
///     Ok(())
/// }
///
pub fn check_sheet_name(name: &str) -> Result<(), XlsxError> {
    let error_message = format!("Invalid Excel worksheet name '{name}'");
    validate_sheetname(name, &error_message)
}

// Internal function to validate worksheet name.
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

// Internal function to validate VBA code names.
pub(crate) fn validate_vba_name(name: &str) -> Result<(), XlsxError> {
    // Check that the  name isn't blank.
    if name.is_empty() {
        return Err(XlsxError::VbaNameError(
            "VBA name cannot be blank".to_string(),
        ));
    }

    // Check that name is <= 31, an Excel limit.
    if name.chars().count() > 31 {
        return Err(XlsxError::VbaNameError(
            "VBA name exceeds Excel limit of 31 characters: {name}".to_string(),
        ));
    }

    // Check for anything other than letters, numbers, and underscores.
    if !name.chars().all(|c| c.is_alphanumeric() || c == '_') {
        return Err(XlsxError::VbaNameError(
            "VBA name contains non-word character: {name}".to_string(),
        ));
    }

    // Check that the name starts with a letter.
    if !name.chars().next().unwrap().is_alphabetic() {
        return Err(XlsxError::VbaNameError(
            "VBA name must start with letter character: {name}".to_string(),
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
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= u16::from(*byte);
    }

    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
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
