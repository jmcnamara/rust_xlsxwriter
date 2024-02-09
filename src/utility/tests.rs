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
}
