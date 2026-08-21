// Utility unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

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
            // Sheet names where the characters before the digits aren't ASCII
            // column letters aren't treated as cell references.
            ("Q.1", "Q.1"),
            ("_1", "_1"),
            ("é1", "é1"),
            ("一1", "一1"),
            ("éa1", "éa1"),
            ("ß1", "ß1"), // Uppercases to "SS" but isn't treated as "SS1".
            ("ﬁ1", "ﬁ1"), // Uppercases to "FI" but isn't treated as "FI1".
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
            assert_eq!(
                exp,
                utility::quote_sheet_name(sheetname),
                "for name '{sheetname}'"
            );
        }
    }

    #[test]
    fn test_is_cell_reference() {
        let tests = vec![
            // Valid names that aren't cell references.
            ("Sales", false),
            ("Table1", false),
            // Outside the Excel column range.
            ("AAAA1", false),
            // Outside the Excel row range.
            ("A1048577", false),
            // A1 style cell references.
            ("A1", true),
            ("a1", true),
            ("A01", true),
            ("XFD1048576", true),
            // R1C1 style cell references, including trailing characters which
            // are ignored by Excel.
            ("R", true),
            ("C", true),
            ("RC", true),
            ("r1c1", true),
            ("R1C1", true),
            ("R1C1x", true),
            ("C1foo", true),
            ("R5x", true),
            ("RC1x", true),
            ("R1c1", true),
            ("rc", true),
            // Potential false matches for cells.
            ("_1", false),
            ("_.1", false),
            ("A.1", false),
            ("A..1", false),
            ("Q.1", false),
            ("é1", false),
            ("É1", false),
            ("Ä1", false),
            ("Ω1", false),
            // Zero row/column references aren't valid cell references.
            ("A0", false),
            ("R0", false),
            ("C0", false),
            ("R0C0", false),
            ("R0C1", false),
            // Outside the Excel row/column range boundaries.
            ("XFE1", false),
            ("R1048577", false),
            ("C1048577", false),
            ("R99999999999999999999999", false),
            // Inside the Excel row/column range boundaries.
            ("R1048576", true),
            ("C16384", true),
            // Names with digits in the middle aren't cell references.
            ("A1A", false),
            ("A1.1", false),
            // Characters that are invalid as lowercase but uppercase into ASCII
            // letters. However, these aren't treated as cell references by Excel.
            ("\u{00DF}1", false), // "ß1" -> "SS".
            ("\u{017F}1", false), // "ſ1" -> "S1".
            ("\u{FB00}1", false), // "ﬀ1" -> "FF1".
            ("\u{FB04}1", false), // "ﬄ1" -> "FFL1".
            ("\u{0130}1", false), // Dotted "İ1".
            ("\u{FF21}1", false), // Fullwidth "Ａ1".
            ("A\u{308}1", false), // NFD "Ä1".
        ];

        for (name, exp) in tests {
            assert_eq!(exp, utility::is_cell_reference(name), "for name '{name}'");
        }
    }

    #[test]
    fn test_check_name() {
        let name_255 = "a".repeat(255);
        let name_256 = "a".repeat(256);

        let tests = vec![
            // Valid names.
            ("Sales", true),
            ("Table1", true),
            ("_name", true),
            ("\\name", true),
            ("My.Name", true),
            ("été", true),
            ("日本語", true),
            // Outside the Excel cell reference row/column ranges.
            ("AAAA1", true),
            ("A1048577", true),
            // At the 255 character limit.
            (name_255.as_str(), true),
            // Excel allows a single underscore or backslash as a name.
            ("_", true),
            ("\\", true),
            // Blank name.
            ("", false),
            // Exceeds the 255 character limit.
            (name_256.as_str(), false),
            // Invalid first characters.
            ("1name", false),
            (".name", false),
            ("?name", false),
            // Underscore and digit are valid.
            ("_1", true),
            // Non-word characters.
            ("name space", false),
            ("name$", false),
            ("name!", false),
            ("name,", false),
            ("name-", false),
            // Cell references.
            ("A1", false),
            ("a1", false),
            ("Z100", false),
            ("XFD1048576", false),
            ("R1C1", false),
            ("r1c1", false),
            // Excel's internally reserved names.
            ("_xlnm.Print_Area", false),
            ("_xlnm._FilterDatabase", false),
            ("_xlnm.Print_Titles", false),
            ("_xlnm.print_area", false),
            ("_XLNM.PRINT_TITLES", false),
            ("_XLNM._FILTERDATABASE", false),
            // Excel's logical constants aren't allowed.
            ("TRUE", false),
            ("FALSE", false),
            ("True", false),
            ("false", false),
            // A backslash followed by a single letter or digit isn't allowed.
            ("\\a", false),
            ("\\z", false),
            ("\\A", false),
            ("\\0", false),
            ("\\9", false),
            ("\\aa", true),
            ("\\11", true),
            // Names in decomposed/NFD form, with combining marks should be
            // valid like their composed/NFC equivalents.
            ("Verk\u{E4}ufe", true),   // NFC "Verkäufe".
            ("Verka\u{308}ufe", true), // NFD "Verkäufe".
            ("e\u{301}", true),        // NFD "é".
            // Characters that uppercase into ASCII letters but aren't treated
            // as cell references.
            ("\u{00DF}1", true), // "ß1" -> "SS1".
            ("\u{017F}1", true), // "ſ1" -> "S1".
            ("\u{FB00}1", true), // "ﬀ1" -> "FF1".
            ("\u{FB04}1", true), // "ﬄ1" -> "FFL1".
            ("\u{0130}1", true), // Dotted "İ1".
            ("\u{FF21}1", true), // Fullwidth "Ａ1".
            ("A\u{308}1", true), // NFD "Ä1".
            // Words that require Unicode marks to spell them.
            ("\u{0E44}\u{0E21}\u{0E48}", true), // Thai, tone mark U+0E48.
            ("\u{0939}\u{093F}\u{0928}\u{094D}\u{0926}\u{0940}", true), // Hindi, virama U+094D.
            ("\u{05D1}\u{0591}", true),         // Hebrew, cantillation mark U+0591.
            // Control characters aren't valid in the target XML.
            ("X\nX", false),
            ("X\tX", false),
        ];

        for (name, exp) in tests {
            assert_eq!(exp, utility::check_name(name).is_ok(), "for name '{name}'");
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
    fn test_splitting_local_name() {
        let tests = vec![
            // Simple unquoted local names.
            ("Sheet1!Name", Some(("Sheet1", "Name"))),
            ("Sheet1!Foo!Bar", Some(("Sheet1", "Foo!Bar"))),
            ("Sheet1!", Some(("Sheet1", ""))),
            // Quoted sheet names, including "!" and escaped quotes.
            ("'Sheet 1'!Name", Some(("'Sheet 1'", "Name"))),
            ("'A!B'!Sales", Some(("'A!B'", "Sales"))),
            ("'It''s'!Sales", Some(("'It''s'", "Sales"))),
            ("'A!''B'!Sales", Some(("'A!''B'", "Sales"))),
            // Global names, without a sheet name part.
            ("Sales", None),
            ("", None),
            // Malformed quoted names. These are treated as Global names and
            // rejected by the `check_name()` validation.
            ("'A!B'Sales", None), // Missing "!" after the quoted part.
            ("'A!B'", None),      // Missing name part.
            ("'A!B!Sales", None), // Unclosed quote.
        ];

        for (name, exp) in tests {
            assert_eq!(exp, utility::split_local_name(name), "for name '{name}'");
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
