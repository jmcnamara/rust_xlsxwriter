// Some utility functions for the `rust_xlsxwriter` module.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::worksheet::ColNum;
use crate::worksheet::RowNum;

// Convert a zero indexed column cell reference to a string.
pub fn col_to_name(col_num: ColNum) -> String {
    let mut col_name = String::new();

    let mut col_num = col_num + 1;

    while col_num > 0 {
        // Set remainder from 1 .. 26
        let mut remainder = col_num % 26;

        if remainder == 0 {
            remainder = 26;
        }

        // Convert the remainder to a character.
        let col_letter = char::from_u32(64u32 + remainder as u32).unwrap();

        // Accumulate the column letters, right to left.
        col_name = format!("{col_letter}{col_name}");

        // Get the next order of magnitude.
        col_num = (col_num - 1) / 26;
    }

    col_name
}

// Convert a column string such as "A" to a zero indexed column reference.
pub fn name_to_col(column: &str) -> ColNum {
    let mut col_num = 0;

    for char in column.chars() {
        col_num = (col_num * 26) + (char as u16 - 'A' as u16 + 1);
    }

    col_num - 1
}

// Convert a zero indexed row and column cell reference to a A1 style string.
pub fn rowcol_to_cell(row_num: RowNum, col_num: ColNum) -> String {
    format!("{}{}", col_to_name(col_num), row_num + 1)
}

// Convert a zero indexed row and column cell reference to an absolute $A$1
// style string.
pub fn rowcol_to_cell_abs(row_num: RowNum, col_num: ColNum) -> String {
    format!("${}${}", col_to_name(col_num), row_num + 1)
}

// Convert zero indexed row and col cell references to a A1:B1 style range string.
pub fn cell_range(
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let range1 = rowcol_to_cell(first_row, first_col);
    let range2 = rowcol_to_cell(last_row, last_col);

    if range1 == range2 {
        range1
    } else {
        format!("{range1}:{range2}")
    }
}

// Convert zero indexed row and col cell references to an absolute $A$1:$B$1
// style range string.
pub fn cell_range_abs(
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let range1 = rowcol_to_cell_abs(first_row, first_col);
    let range2 = rowcol_to_cell_abs(last_row, last_col);

    if range1 == range2 {
        range1
    } else {
        format!("{range1}:{range2}")
    }
}

// Convert zero indexed row and col cell references to a chart absolute
// Sheet1!$A$1:$B$1 style range string.
pub fn chart_range_abs(
    sheet_name: &str,
    first_row: RowNum,
    first_col: ColNum,
    last_row: RowNum,
    last_col: ColNum,
) -> String {
    let sheet_name = quote_sheetname(sheet_name);
    let range1 = rowcol_to_cell_abs(first_row, first_col);
    let range2 = rowcol_to_cell_abs(last_row, last_col);

    if range1 == range2 {
        format!("{sheet_name}!{range1}")
    } else {
        format!("{sheet_name}!{range1}:{range2}")
    }
}

// Create a quoted version of a worksheet name. Excel single quotes worksheet
// names that contain spaces and some other characters.
pub fn quote_sheetname(sheetname: &str) -> String {
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

// Get the pixel width of a string based on character widths taken from Excel.
// Non-ascii characters are given a default width of 8 pixels.
#[allow(clippy::match_same_arms)]
pub fn pixel_width(string: &str) -> u16 {
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
        hash ^= *byte as u16;
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

//
// Tests.
//
#[cfg(test)]
mod tests {

    use crate::utility;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_hash_password() {
        let tests = vec![
            ("", "0000"),
            ("password", "83AF"),
            ("This is a longer phrase", "D14E"),
            ("0", "CE2A"),
            ("01", "CEED"),
            ("012", "CF7C"),
            ("0123", "CC4B"),
            ("01234", "CACA"),
            ("012345", "C789"),
            ("0123456", "DC88"),
            ("01234567", "EB87"),
            ("012345678", "9B86"),
            ("0123456789", "FF84"),
            ("01234567890", "FF86"),
            ("012345678901", "EF87"),
            ("0123456789012", "AF8A"),
            ("01234567890123", "EF90"),
            ("012345678901234", "EFA5"),
            ("0123456789012345", "EFD0"),
            ("01234567890123456", "EF09"),
            ("012345678901234567", "EEB2"),
            ("0123456789012345678", "ED33"),
            ("01234567890123456789", "EA14"),
            ("012345678901234567890", "E615"),
            ("0123456789012345678901", "FE96"),
            ("01234567890123456789012", "CC97"),
            ("012345678901234567890123", "AA98"),
            ("0123456789012345678901234", "FA98"),
            ("01234567890123456789012345", "D298"),
            ("0123456789012345678901234567890", "D2D3"),
        ];

        for (string, exp) in tests {
            let got = format!("{:04X}", utility::hash_password(string));
            assert_eq!(exp, got);
        }
    }

    #[test]
    fn test_col_to_name() {
        let tests = vec![
            (0, "A"),
            (1, "B"),
            (2, "C"),
            (9, "J"),
            (24, "Y"),
            (25, "Z"),
            (26, "AA"),
            (254, "IU"),
            (255, "IV"),
            (256, "IW"),
            (16383, "XFD"),
            (16384, "XFE"),
        ];

        for (col_num, col_string) in tests {
            assert_eq!(col_string, utility::col_to_name(col_num));
        }
    }

    #[test]
    fn test_name_to_col() {
        let tests = vec![
            (0, "A"),
            (1, "B"),
            (2, "C"),
            (9, "J"),
            (24, "Y"),
            (25, "Z"),
            (26, "AA"),
            (254, "IU"),
            (255, "IV"),
            (256, "IW"),
            (16383, "XFD"),
            (16384, "XFE"),
        ];

        for (col_num, col_string) in tests {
            assert_eq!(col_num, utility::name_to_col(col_string));
        }
    }

    #[test]
    fn test_row_col_to_cell() {
        let tests = vec![
            (0, 0, "A1"),
            (0, 1, "B1"),
            (0, 2, "C1"),
            (0, 9, "J1"),
            (1, 0, "A2"),
            (2, 0, "A3"),
            (9, 0, "A10"),
            (1, 24, "Y2"),
            (7, 25, "Z8"),
            (9, 26, "AA10"),
            (1, 254, "IU2"),
            (1, 255, "IV2"),
            (1, 256, "IW2"),
            (0, 16383, "XFD1"),
            (1048576, 16384, "XFE1048577"),
        ];

        for (row_num, col_num, cell_string) in tests {
            assert_eq!(cell_string, utility::rowcol_to_cell(row_num, col_num));
        }
    }

    #[test]
    fn test_cell_range() {
        let tests = vec![
            (0, 0, 9, 0, "A1:A10"),
            (1, 2, 8, 2, "C2:C9"),
            (0, 0, 3, 4, "A1:E4"),
            (0, 0, 0, 0, "A1"),
            (0, 0, 0, 1, "A1:B1"),
            (0, 2, 0, 9, "C1:J1"),
            (1, 0, 2, 0, "A2:A3"),
            (9, 0, 1, 24, "A10:Y2"),
            (7, 25, 9, 26, "Z8:AA10"),
            (1, 254, 1, 255, "IU2:IV2"),
            (1, 256, 0, 16383, "IW2:XFD1"),
            (0, 0, 1048576, 16384, "A1:XFE1048577"),
        ];

        for (start_row, start_col, end_row, end_col, cell_range) in tests {
            assert_eq!(
                cell_range,
                utility::cell_range(start_row, start_col, end_row, end_col)
            );
        }
    }

    #[test]
    fn test_quote_sheetname() {
        let tests = vec![
            ("Sheet1", "Sheet1"),
            ("Sheet.2", "Sheet.2"),
            ("Sheet_3", "Sheet_3"),
            ("'Sheet4'", "'Sheet4'"),
            ("'Sheet 5'", "Sheet 5"),
            ("'Sheet!6'", "Sheet!6"),
            ("'Sheet''7'", "Sheet'7"),
            (
                "'a''''''''''''''''''''''''''''''''''''''''''''''''''''''''''b'",
                "a'''''''''''''''''''''''''''''b",
            ),
        ];

        for (exp, sheetname) in tests {
            assert_eq!(exp, utility::quote_sheetname(sheetname));
        }
    }

    #[test]
    fn test_pixel_width() {
        let tests = vec![
            (" ", 3),
            ("!", 5),
            ("\"", 6),
            ("#", 7),
            ("$", 7),
            ("%", 11),
            ("&", 10),
            ("'", 3),
            ("(", 5),
            (")", 5),
            ("*", 7),
            ("+", 7),
            (",", 4),
            ("-", 5),
            (".", 4),
            ("/", 6),
            ("0", 7),
            ("1", 7),
            ("2", 7),
            ("3", 7),
            ("4", 7),
            ("5", 7),
            ("6", 7),
            ("7", 7),
            ("8", 7),
            ("9", 7),
            (":", 4),
            (";", 4),
            ("<", 7),
            ("=", 7),
            (">", 7),
            ("?", 7),
            ("@", 13),
            ("A", 9),
            ("B", 8),
            ("C", 8),
            ("D", 9),
            ("E", 7),
            ("F", 7),
            ("G", 9),
            ("H", 9),
            ("I", 4),
            ("J", 5),
            ("K", 8),
            ("L", 6),
            ("M", 12),
            ("N", 10),
            ("O", 10),
            ("P", 8),
            ("Q", 10),
            ("R", 8),
            ("S", 7),
            ("T", 7),
            ("U", 9),
            ("V", 9),
            ("W", 13),
            ("X", 8),
            ("Y", 7),
            ("Z", 7),
            ("[", 5),
            ("\\", 6),
            ("]", 5),
            ("^", 7),
            ("_", 7),
            ("`", 4),
            ("a", 7),
            ("b", 8),
            ("c", 6),
            ("d", 8),
            ("e", 8),
            ("f", 5),
            ("g", 7),
            ("h", 8),
            ("i", 4),
            ("j", 4),
            ("k", 7),
            ("l", 4),
            ("m", 12),
            ("n", 8),
            ("o", 8),
            ("p", 8),
            ("q", 8),
            ("r", 5),
            ("s", 6),
            ("t", 5),
            ("u", 8),
            ("v", 7),
            ("w", 11),
            ("x", 7),
            ("y", 7),
            ("z", 6),
            ("{", 5),
            ("|", 7),
            ("}", 5),
            ("~", 7),
            ("é", 8),
            ("éé", 16),
            ("ABC", 25),
            ("Hello", 33),
            ("12345", 35),
        ];

        for (string, exp) in tests {
            assert_eq!(exp, utility::pixel_width(string));
        }
    }
}
