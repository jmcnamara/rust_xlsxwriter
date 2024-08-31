// Utility unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod utility_tests {

    use crate::{utility, XlsxError};
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
            assert_eq!(col_string, utility::column_number_to_name(col_num));
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
            assert_eq!(col_num, utility::column_name_to_number(col_string));
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
            assert_eq!(cell_string, utility::row_col_to_cell(row_num, col_num));
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
    // The following unquoted and quoted sheet names were extracted from
    // Excel files.
    fn test_quote_sheetname() {
        let tests = vec![
            // A sheetname that is already quoted.
            ("'Sheet 1'", "'Sheet 1'"),
            // ----------------------------------------------------------------
            // Rule 1.
            // ----------------------------------------------------------------
            // Some simple variants on standard sheet names.
            ("Sheet1", "Sheet1"),
            ("Sheet.1", "Sheet.1"),
            ("Sheet_1", "Sheet_1"),
            ("Sheet-1", "'Sheet-1'"),
            ("Sheet 1", "'Sheet 1'"),
            ("Sheet#1", "'Sheet#1'"),
            // Sheetnames with single quotes.
            ("Sheet'1", "'Sheet''1'"),
            ("Sheet''1", "'Sheet''''1'"),
            // Single special chars that are unquoted in sheetnames. These are
            // variants of the first char rule.
            ("_", "_"),
            (".", "'.'"),
            // White space only.
            (" ", "' '"),
            ("    ", "'    '"),
            // Sheetnames with unicode or emojis.
            ("été", "été"),
            ("mangé", "mangé"),
            ("Sheet©", "'Sheet©'"),
            ("Sheet😀", "Sheet😀"),
            ("Sheet🤌1", "Sheet🤌1"),
            ("Sheet⟦1", "'Sheet⟦1'"), // Unicode punctuation.
            ("Sheet᠅1", "'Sheet᠅1'"), // Unicode punctuation.
            // ----------------------------------------------------------------
            // Rule 2.
            // ----------------------------------------------------------------
            // Sheetnames starting with non-word characters.
            ("_Sheet1", "_Sheet1"),
            (".Sheet1", "'.Sheet1'"),
            ("1Sheet1", "'1Sheet1'"),
            ("-Sheet1", "'-Sheet1'"),
            ("#Sheet1", "'#Sheet1'"),
            ("©Sheet", "'©Sheet'"),
            ("😀Sheet", "'😀Sheet'"),
            ("🤌Sheet", "'🤌Sheet'"),
            // Sheetnames that are digits only also start with a non word char.
            ("1", "'1'"),
            ("2", "'2'"),
            ("1234", "'1234'"),
            ("12345678", "'12345678'"),
            // ----------------------------------------------------------------
            // Rule 3.
            // ----------------------------------------------------------------
            // Worksheet names that look like A1 style references (with the
            // row/column number in the Excel allowable range). These are case
            // insensitive.
            ("A0", "A0"),
            ("A1", "'A1'"),
            ("a1", "'a1'"),
            ("XFD", "XFD"),
            ("xfd", "xfd"),
            ("XFE1", "XFE1"),
            ("ZZZ1", "ZZZ1"),
            ("XFD1", "'XFD1'"),
            ("xfd1", "'xfd1'"),
            ("B1048577", "B1048577"),
            ("A1048577", "A1048577"),
            ("A1048576", "'A1048576'"),
            ("B1048576", "'B1048576'"),
            ("B1048576a", "B1048576a"),
            ("XFD048576", "'XFD048576'"),
            ("XFD1048576", "'XFD1048576'"),
            ("XFD01048577", "XFD01048577"),
            ("XFD01048576", "'XFD01048576'"),
            ("A123456789012345678901", "A123456789012345678901"), // Exceeds u64.
            // ----------------------------------------------------------------
            // Rule 4.
            // ----------------------------------------------------------------
            // Sheet names that *start* with RC style references (with the
            // row/column number in the Excel allowable range). These are case
            // insensitive.
            ("A", "A"),
            ("B", "B"),
            ("D", "D"),
            ("Q", "Q"),
            ("S", "S"),
            ("c", "'c'"),
            ("C", "'C'"),
            ("CR", "CR"),
            ("CZ", "CZ"),
            ("r", "'r'"),
            ("R", "'R'"),
            ("C8", "'C8'"),
            ("rc", "'rc'"),
            ("RC", "'RC'"),
            ("RCZ", "RCZ"),
            ("RRC", "RRC"),
            ("R0C0", "R0C0"),
            ("R4C", "'R4C'"),
            ("R5C", "'R5C'"),
            ("rc2", "'rc2'"),
            ("RC2", "'RC2'"),
            ("RC8", "'RC8'"),
            ("bR1C1", "bR1C1"),
            ("R1C1", "'R1C1'"),
            ("r1c2", "'r1c2'"),
            ("rc2z", "'rc2z'"),
            ("bR1C1b", "bR1C1b"),
            ("R1C1b", "'R1C1b'"),
            ("R1C1R", "'R1C1R'"),
            ("C16384", "'C16384'"),
            ("C16385", "'C16385'"),
            ("C16385Z", "C16385Z"),
            ("C16386", "'C16386'"),
            ("C16384Z", "'C16384Z'"),
            ("PC16384Z", "PC16384Z"),
            ("RC16383", "'RC16383'"),
            ("RC16385Z", "RC16385Z"),
            ("R1048576", "'R1048576'"),
            ("R1048577C", "R1048577C"),
            ("R1C16384", "'R1C16384'"),
            ("R1C16385", "'R1C16385'"),
            ("RC16384Z", "'RC16384Z'"),
            ("R1048576C", "'R1048576C'"),
            ("R1048577C1", "R1048577C1"),
            ("R1C16384Z", "'R1C16384Z'"),
            ("R1048575C1", "'R1048575C1'"),
            ("R1048576C1", "'R1048576C1'"),
            ("R1048577C16384", "R1048577C16384"),
            ("R1048576C16384", "'R1048576C16384'"),
            ("R1048576C16385", "'R1048576C16385'"),
            ("ZR1048576C16384", "ZR1048576C16384"),
            ("C123456789012345678901Z", "C123456789012345678901Z"), // Exceeds u64.
            ("R123456789012345678901Z", "R123456789012345678901Z"), // Exceeds u64.
        ];

        for (sheetname, exp) in tests {
            assert_eq!(exp, utility::quote_sheetname(sheetname));
        }
    }

    #[test]
    fn test_unquote_sheetname() {
        let tests = vec![
            ("Sheet1", "Sheet1"),
            ("'Sheet2'", "Sheet2"),
            ("'Sheet''3'", "Sheet'3"),
            ("'Sheet''''4'", "Sheet''4"),
            (
                "'a''''''''''''''''''''''''''''''''''''''''''''''''''''''''''b'",
                "a'''''''''''''''''''''''''''''b",
            ),
        ];
        for (sheetname, exp) in tests {
            assert_eq!(exp, utility::unquote_sheetname(sheetname));
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

    #[test]
    fn check_invalid_worksheet_names() {
        let result = utility::check_sheet_name("");
        assert!(matches!(result, Err(XlsxError::SheetnameCannotBeBlank(_))));

        let name = "name_that_is_longer_than_thirty_one_characters";
        let result = utility::check_sheet_name(name);
        assert!(matches!(result, Err(XlsxError::SheetnameLengthExceeded(_))));

        let name = "name_with_special_character_[";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_]";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_:";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_*";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_?";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_/";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_\\";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "'start with apostrophe";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameStartsOrEndsWithApostrophe(_))
        ));

        let name = "end with apostrophe'";
        let result = utility::check_sheet_name(name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameStartsOrEndsWithApostrophe(_))
        ));
    }

    #[test]
    fn check_invalid_vba_names() {
        let result = utility::validate_vba_name("ValidName");
        assert!(matches!(result, Ok(())));

        let result = utility::validate_vba_name("Alphanumeric_characters_123");
        assert!(matches!(result, Ok(())));

        let result = utility::validate_vba_name("");
        assert!(matches!(result, Err(XlsxError::VbaNameError(_))));

        let name = "name_that_is_longer_than_thirty_one_characters";
        let result = utility::validate_vba_name(name);
        assert!(matches!(result, Err(XlsxError::VbaNameError(_))));

        let name = "name_with_non_word_character_*";
        let result = utility::validate_vba_name(name);
        assert!(matches!(result, Err(XlsxError::VbaNameError(_))));

        let name = "1name_starts_with_non_letter_char";
        let result = utility::validate_vba_name(name);
        assert!(matches!(result, Err(XlsxError::VbaNameError(_))));

        let name = "_name_starts_with_non_letter_char";
        let result = utility::validate_vba_name(name);
        assert!(matches!(result, Err(XlsxError::VbaNameError(_))));
    }

    #[test]
    fn check_is_valid_range() {
        assert_eq!(true, utility::is_valid_range("A1"));
        assert_eq!(true, utility::is_valid_range("A1:B3"));

        assert_eq!(false, utility::is_valid_range(""));
        assert_eq!(false, utility::is_valid_range("1A"));
        assert_eq!(false, utility::is_valid_range("a1"));
        assert_eq!(false, utility::is_valid_range("1:3"));
    }
}
