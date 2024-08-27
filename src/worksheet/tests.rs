// worksheet unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod worksheet_tests {

    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
    use crate::XlsxError;
    use pretty_assertions::assert_eq;
    use std::collections::HashMap;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.assemble_xml_file();

        let got = worksheet.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn verify_header_footer_images() {
        let strings = [
            ("", HeaderImagePosition::Left, false),
            ("&L&G", HeaderImagePosition::Left, true),
            ("&R&[Picture]", HeaderImagePosition::Right, true),
            ("&C&[Picture]", HeaderImagePosition::Center, true),
            ("&R&[Picture]", HeaderImagePosition::Left, false),
            ("&L&[Picture]&C&[Picture]", HeaderImagePosition::Left, true),
            (
                "&L&[Picture]&C&[Picture]",
                HeaderImagePosition::Center,
                true,
            ),
            (
                "&L&[Picture]&C&[Picture]",
                HeaderImagePosition::Right,
                false,
            ),
            ("&L&[Picture]", HeaderImagePosition::Left, true),
            (
                "&L&[Picture]&C&[Picture]&R&[Picture]",
                HeaderImagePosition::Left,
                true,
            ),
            (
                "&C&[Picture]&L&G&R&[Picture]",
                HeaderImagePosition::Left,
                true,
            ),
        ];

        for (string, position, exp) in strings {
            assert_eq!(
                exp,
                Worksheet::verify_header_footer_image(string, &position)
            );
        }
    }

    #[test]
    #[cfg(feature = "serde")]
    fn get_serialize_dimensions() {
        let mut worksheet = Worksheet::new();

        #[derive(Serialize)]
        struct MyStruct {
            column1: u8,
            column2: u8,
            column3: u8,
            column4: u8,
        }

        let data = MyStruct {
            column1: 1,
            column2: 2,
            column3: 3,
            column4: 4,
        };

        worksheet.serialize_headers(2, 2, &data).unwrap();

        for _ in 1..=10 {
            worksheet.serialize(&data).unwrap();
        }

        let result = worksheet.get_serialize_dimensions("MyStruct").unwrap();
        assert_eq!((2, 2, 12, 5), result);

        let result = worksheet
            .get_serialize_column_dimensions("MyStruct", "column1")
            .unwrap();
        assert_eq!((2, 2, 12, 2), result);

        let result = worksheet
            .get_serialize_column_dimensions("MyStruct", "column4")
            .unwrap();
        assert_eq!((2, 5, 12, 5), result);

        let result = worksheet.get_serialize_dimensions("Doesn't exist");
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));

        let result = worksheet.get_serialize_column_dimensions("Doesn't exist", "column1");
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));

        let result = worksheet.get_serialize_column_dimensions("MyStruct", "Doesn't exist");
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));
    }

    #[test]
    fn row_matches_list_filter_blanks() {
        let mut worksheet = Worksheet::new();
        let bold = Format::new().set_bold();

        worksheet.write_string(0, 0, "Header").unwrap();
        worksheet.write_string(1, 0, "").unwrap();
        worksheet.write_string(2, 0, " ").unwrap();
        worksheet.write_string(3, 0, "  ").unwrap();
        worksheet.write_string_with_format(4, 0, "", &bold).unwrap();

        let filter_condition = FilterCondition::new().add_list_blanks_filter();

        assert!(!worksheet.row_matches_list_filter(0, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(1, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(2, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(3, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(4, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(5, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(7, 7, &filter_condition));
    }

    #[test]
    fn row_matches_list_filter_strings() {
        let mut worksheet = Worksheet::new();
        worksheet.write_string(0, 0, "Header").unwrap();
        worksheet.write_string(1, 0, "South").unwrap();
        worksheet.write_string(2, 0, "south").unwrap();
        worksheet.write_string(3, 0, "SOUTH").unwrap();
        worksheet.write_string(4, 0, "South ").unwrap();
        worksheet.write_string(5, 0, " South").unwrap();
        worksheet.write_string(6, 0, " South ").unwrap();
        worksheet.write_string(7, 0, "Mouth").unwrap();

        let filter_condition = FilterCondition::new().add_list_filter("South");

        assert!(worksheet.row_matches_list_filter(1, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(2, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(3, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(4, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(5, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(6, 0, &filter_condition));
        assert!(!worksheet.row_matches_list_filter(7, 0, &filter_condition));
    }

    #[test]
    fn row_matches_list_filter_numbers() {
        let mut worksheet = Worksheet::new();

        worksheet.write_string(0, 0, "Header").unwrap();
        worksheet.write_number(1, 0, 1000).unwrap();
        worksheet.write_number(2, 0, 1000.0).unwrap();
        worksheet.write_string(3, 0, "1000").unwrap();
        worksheet.write_string(4, 0, " 1000 ").unwrap();
        worksheet.write_number(5, 0, 2000).unwrap();

        let filter_condition = FilterCondition::new().add_list_filter(1000);

        assert!(worksheet.row_matches_list_filter(1, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(2, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(3, 0, &filter_condition));
        assert!(worksheet.row_matches_list_filter(4, 0, &filter_condition));
        assert!(!worksheet.row_matches_list_filter(5, 0, &filter_condition));
    }

    #[test]
    fn process_pagebreaks() {
        let mut worksheet = Worksheet::new();

        // Test removing duplicates.
        let got = Worksheet::process_pagebreaks(&[1, 1, 1, 1]).unwrap();
        assert_eq!(vec![1], got);

        // Test removing 0.
        let got = Worksheet::process_pagebreaks(&[0, 1, 2, 3, 4]).unwrap();
        assert_eq!(vec![1, 2, 3, 4], got);

        // Test sort order.
        let got = Worksheet::process_pagebreaks(&[1, 12, 2, 13, 3, 4]).unwrap();
        assert_eq!(vec![1, 2, 3, 4, 12, 13], got);

        // Exceed the number of allow breaks.
        let breaks = (1u32..=1024).collect::<Vec<u32>>();
        let result = Worksheet::process_pagebreaks(&breaks);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));

        // Test row and column limits.
        let result = worksheet.set_page_breaks(&[ROW_MAX]);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_vertical_page_breaks(&[COL_MAX as u32]);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));
    }

    #[test]
    fn set_header_image() {
        let mut worksheet = Worksheet::new();

        let image = Image::new("tests/input/images/red.jpg").unwrap();
        worksheet.set_header("&R&G");

        // Test inserting an image without a matching header position.
        let result = worksheet.set_header_image(&image, HeaderImagePosition::Left);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));
    }

    #[test]
    fn rich_string() {
        let mut worksheet = Worksheet::new();

        // Test an empty array.
        let segments = [];
        let result = worksheet.write_rich_string(0, 0, &segments);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));

        // Test an empty string.
        let default = Format::default();
        let segments = [(&default, "")];
        let result = worksheet.write_rich_string(0, 0, &segments);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));
    }

    #[test]
    fn test_calculate_spans_1() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (0..17).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:16".to_string()), (1, "17:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_2() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (1..18).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:15".to_string()), (1, "16:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_3() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (2..19).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:14".to_string()), (1, "15:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_4() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (3..20).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:13".to_string()), (1, "14:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_5() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (4..21).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:12".to_string()), (1, "13:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_6() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (5..22).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:11".to_string()), (1, "12:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_7() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (6..23).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:10".to_string()), (1, "11:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_8() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (7..24).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:9".to_string()), (1, "10:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_9() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (8..25).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:8".to_string()), (1, "9:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_10() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (9..26).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:7".to_string()), (1, "8:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_11() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (10..27).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:6".to_string()), (1, "7:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_12() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (11..28).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:5".to_string()), (1, "6:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_13() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (12..29).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:4".to_string()), (1, "5:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_14() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (13..30).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:3".to_string()), (1, "4:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_15() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (14..31).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:2".to_string()), (1, "3:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_16() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (15..32).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:1".to_string()), (1, "2:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_17() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (16..33).enumerate() {
            worksheet
                .write_number(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(1, "1:16".to_string()), (2, "17:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn test_calculate_spans_18() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (16..33).enumerate() {
            worksheet
                .write_number(row_num, (col_num + 1) as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(1, "2:17".to_string()), (2, "18:18".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(expected, got);
    }

    #[test]
    fn check_invalid_worksheet_names() {
        let mut worksheet = Worksheet::new();

        let result = worksheet.set_name("");
        assert!(matches!(result, Err(XlsxError::SheetnameCannotBeBlank(_))));

        let name = "name_that_is_longer_than_thirty_one_characters".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(result, Err(XlsxError::SheetnameLengthExceeded(_))));

        let name = "name_with_special_character_[".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_]".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_:".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_*".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_?".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_/".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "name_with_special_character_\\".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameContainsInvalidCharacter(_))
        ));

        let name = "'start with apostrophe".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameStartsOrEndsWithApostrophe(_))
        ));

        let name = "end with apostrophe'".to_string();
        let result = worksheet.set_name(&name);
        assert!(matches!(
            result,
            Err(XlsxError::SheetnameStartsOrEndsWithApostrophe(_))
        ));
    }

    #[test]
    fn get_name() {
        let mut worksheet = Worksheet::new();

        let got = worksheet.name();
        assert_eq!("", got);

        let exp = "Sheet1";
        worksheet.set_name(exp).unwrap();
        let got = worksheet.name();
        assert_eq!(exp, got);
    }

    #[test]
    fn merge_range() {
        let mut worksheet = Worksheet::new();
        let format = Format::default();

        // Test single merge cell.
        let result = worksheet.merge_range(1, 1, 1, 1, "Foo", &format);
        assert!(matches!(result, Err(XlsxError::MergeRangeSingleCell)));

        // Test for overlap.
        let _worksheet = worksheet.merge_range(1, 1, 20, 20, "Foo", &format);
        let result = worksheet.merge_range(2, 2, 3, 3, "Foo", &format);
        assert!(matches!(result, Err(XlsxError::MergeRangeOverlaps(_, _))));

        // Test out of range value.
        let result = worksheet.merge_range(ROW_MAX, 1, 1, 1, "Foo", &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        // Test out reversed values
        let result = worksheet.merge_range(5, 1, 1, 1, "Foo", &format);
        assert!(matches!(result, Err(XlsxError::RowColumnOrderError)));
    }

    #[test]
    fn check_dimensions() {
        let mut worksheet = Worksheet::new();
        let format = Format::default();

        assert!(!worksheet.check_dimensions(ROW_MAX, 0));
        assert!(!worksheet.check_dimensions(0, COL_MAX));

        let result = worksheet.write_string_with_format(ROW_MAX, 0, "Foo", &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.write_string(ROW_MAX, 0, "Foo");
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.write_number_with_format(ROW_MAX, 0, 0, &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.write_number(ROW_MAX, 0, 0);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_row_height(ROW_MAX, 20);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_row_height_pixels(ROW_MAX, 20);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_row_format(ROW_MAX, &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_column_width(COL_MAX, 20);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_column_width_pixels(COL_MAX, 20);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_column_format(COL_MAX, &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));
    }

    #[test]
    fn long_string() {
        let mut worksheet = Worksheet::new();
        let chars: [u8; 32_768] = [64; 32_768];
        let long_string = std::str::from_utf8(&chars);

        let result = worksheet.write_string(0, 0, long_string.unwrap());
        assert!(matches!(result, Err(XlsxError::MaxStringLengthExceeded)));
    }
}
