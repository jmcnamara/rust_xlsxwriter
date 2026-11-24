// worksheet unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod worksheet_tests {

    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
    use crate::{
        DataValidation, FilterCondition, FilterCriteria, Format, Sparkline, Table, TableColumn,
        Url, XlsxError,
    };
    use pretty_assertions::assert_eq;
    use std::collections::HashMap;
    use std::io::Cursor;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
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
    fn test_data_validation_with_escaped_strings() {
        let mut worksheet = Worksheet::new();

        let data_validation = crate::DataValidation::new()
            .allow_whole_number(crate::DataValidationRule::Between(1, 5))
            .set_input_title("a_x0041_b")
            .unwrap()
            .set_input_message("a_x0041_b")
            .unwrap()
            .set_error_title("a_x0041_b")
            .unwrap()
            .set_error_message("a_x0041_b")
            .unwrap();

        worksheet
            .add_data_validation(0, 0, 0, 0, &data_validation)
            .unwrap();

        worksheet.write_data_validations();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <dataValidations count="1">
              <dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="a_x005f_x0041_b" error="a_x005f_x0041_b" promptTitle="a_x005f_x0041_b" prompt="a_x005f_x0041_b" sqref="A1">
                <formula1>1</formula1>
                <formula2>5</formula2>
              </dataValidation>
            </dataValidations>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_hyperlink_with_escaped_strings() {
        let mut worksheet = Worksheet::new();

        let url = crate::Url::new("internal:Sheet1!A1")
            .set_text("a_x0041_b")
            .set_tip("a_x0041_b");
        worksheet.write_url(0, 0, &url).unwrap();

        worksheet.write_hyperlinks();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <hyperlinks>
              <hyperlink ref="A1" location="Sheet1!A1" tooltip="a_x005f_x0041_b" display="a_x005f_x0041_b"/>
            </hyperlinks>
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
    fn test_merge_range_with_rejected_overlap() {
        let mut worksheet = Worksheet::new();
        let format = Format::default();

        // B1:B2.
        worksheet.merge_range(0, 1, 1, 1, "Foo", &format).unwrap();

        // A partially overlapping range, A1:B2, shouldn't change the worksheet.
        let result = worksheet.merge_range(0, 0, 1, 1, "Bar", &format);
        assert!(matches!(result, Err(XlsxError::MergeRangeOverlaps(_, _))));
        assert!(!worksheet.data_table[&0].contains_key(&0));
        assert_eq!(worksheet.merged_cells.len(), 2);

        // The cells that didn't overlap, A1:A2, should still be available.
        worksheet.merge_range(0, 0, 1, 0, "Bar", &format).unwrap();
    }

    #[test]
    fn test_add_table_with_rejected_overlap() {
        let mut worksheet = Worksheet::new();
        let columns = [
            TableColumn::new().set_header("Foo"),
            TableColumn::new().set_header("Bar"),
        ];
        let table = Table::new();
        let table2 = Table::new().set_columns(&columns);

        // B1:B3.
        worksheet.add_table(0, 1, 2, 1, &table).unwrap();

        // A partially overlapping range, A1:B3, shouldn't change the worksheet.
        let result = worksheet.add_table(0, 0, 2, 1, &table2);
        assert!(matches!(result, Err(XlsxError::TableRangeOverlaps(_, _))));
        assert!(!worksheet.data_table[&0].contains_key(&0));
        assert_eq!(worksheet.table_cells.len(), 3);
        assert_eq!(worksheet.tables.len(), 1);

        // The cells that didn't overlap, A1:A3, should still be available.
        worksheet.add_table(0, 0, 2, 0, &table).unwrap();
    }

    #[test]
    fn test_add_table_with_rejected_autofilter_overlap() {
        let mut worksheet = Worksheet::new();
        let table = Table::new();

        // A1:A6.
        worksheet.autofilter(0, 0, 5, 0).unwrap();

        // A table, A1:B3, that overlaps the autofilter shouldn't change the
        // worksheet.
        let result = worksheet.add_table(0, 0, 2, 1, &table);
        assert!(matches!(
            result,
            Err(XlsxError::AutofilterRangeOverlaps(_, _))
        ));
        assert!(worksheet.data_table.is_empty());
        assert!(worksheet.table_cells.is_empty());
        assert_eq!(worksheet.cells_with_autofilter.len(), 1);
        assert!(worksheet.tables.is_empty());

        // A table outside the autofilter, B2:B4, should still be allowed.
        worksheet.add_table(1, 1, 3, 1, &table).unwrap();
    }

    #[test]
    fn test_autofilter_with_rejected_table_overlap() {
        let mut worksheet = Worksheet::new();
        let table = Table::new();

        // B1:B3.
        worksheet.add_table(0, 1, 2, 1, &table).unwrap();

        // An autofilter, A1:B6, that overlaps the table shouldn't change the
        // worksheet.
        let result = worksheet.autofilter(0, 0, 5, 1);
        assert!(matches!(
            result,
            Err(XlsxError::AutofilterRangeOverlaps(_, _))
        ));
        assert!(!worksheet.autofilter_defined_name.in_use);
        assert!(worksheet.autofilter_area.is_empty());
        assert_eq!(worksheet.cells_with_autofilter.len(), 1);
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

        // Excel's limit is in UTF-16 units, so these emojis count as 2.
        let long_string = "\u{1F600}".repeat(16_384);
        let result = worksheet.write_string(0, 0, &long_string);
        assert!(matches!(result, Err(XlsxError::MaxStringLengthExceeded)));

        let long_string = "\u{1F600}".repeat(16_383) + "a";
        let result = worksheet.write_string(0, 0, &long_string);
        assert!(result.is_ok());
    }

    #[test]
    fn ignored_cells_reuse() {
        let mut worksheet = Worksheet::new();

        // Insert a rule.
        let result = worksheet.ignore_error(2, 1, IgnoreError::NumberStoredAsText);
        assert!(result.is_ok());

        // Test reuse of the same cell.
        let result = worksheet.ignore_error(2, 1, IgnoreError::NumberStoredAsText);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));
    }

    #[test]
    fn assert_sync() {
        const _: () = {
            const fn assert_worksheet_sync<S: Sync>() {}
            assert_worksheet_sync::<Worksheet>();
        };
    }

    #[test]
    fn test_sheet_views2_1() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_freeze_panes(1, 0).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
                    <selection pane="bottomLeft"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views2_2() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_freeze_panes(0, 1).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
                    <selection pane="topRight"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views2_3() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_freeze_panes(1, 1).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
                    <selection pane="topRight" activeCell="B1" sqref="B1"/>
                    <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
                    <selection pane="bottomRight"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views2_4() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_freeze_panes(3, 6).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/>
                    <selection pane="topRight" activeCell="G1" sqref="G1"/>
                    <selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
                    <selection pane="bottomRight"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views6_1() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_selection(1, 0, 1, 0).unwrap();
        worksheet.set_freeze_panes(1, 0).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
                    <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views6_2() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_selection(0, 1, 0, 1).unwrap();
        worksheet.set_freeze_panes(0, 1).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
                    <selection pane="topRight" activeCell="B1" sqref="B1"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views6_3() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_selection(3, 6, 3, 6).unwrap();
        worksheet.set_freeze_panes(3, 6).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/>
                    <selection pane="topRight" activeCell="G1" sqref="G1"/>
                    <selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
                    <selection pane="bottomRight" activeCell="G4" sqref="G4"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views6_4() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_selection(4, 8, 4, 8).unwrap();
        worksheet.set_freeze_panes(3, 6).unwrap();

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/>
                    <selection pane="topRight" activeCell="G1" sqref="G1"/>
                    <selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
                    <selection pane="bottomRight" activeCell="I5" sqref="I5"/>
                </sheetView>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sheet_views_hidden_zeroes() {
        let mut worksheet = Worksheet {
            selected: true,
            ..Default::default()
        };

        worksheet.set_hidden_zeroes(true);

        worksheet.write_sheet_views();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <sheetViews>
                <sheetView showZeros="0" tabSelected="1" workbookViewId="0"/>
            </sheetViews>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_formulas_with_escapes() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        Worksheet::write_formula_cell(&mut writer, 1, "A", r#""a_x0041_b""#, 0, "a_x0041_b");
        Worksheet::write_array_formula_cell(
            &mut writer,
            2,
            "A",
            r#""a_x0041_b""#,
            0,
            "a_x0041_b",
            false,
            "A2",
        );

        let got = xmlwriter::cursor_to_str(&writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <c r="A1" t="str"><f>"a_x005F_x0041_b"</f><v>a_x005F_x0041_b</v></c>
            <c r="A2" t="str"><f t="array" ref="A2">"a_x005F_x0041_b"</f><v>a_x005F_x0041_b</v></c>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_rich_string_with_escapes() {
        let bold = Format::new().set_bold();
        let default = Format::default();

        // Excel escapes each run separately.
        let segments = [(&bold, "a_x0041_b"), (&default, "a_x00"), (&bold, "41_b")];

        let (got, _) = Worksheet::get_rich_string(&segments).unwrap();
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
            <r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>a_x005F_x0041_b</t></r>
            <r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>a_x00</t></r>
            <r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>41_b</t></r>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_header_footer_with_escapes() {
        let mut worksheet = Worksheet::new();
        worksheet.set_header("&Ca_x0041_b");
        worksheet.set_footer("&Ca_x0041_b");

        worksheet.write_header_footer();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <headerFooter>
              <oddHeader>&amp;Ca_x005F_x0041_b</oddHeader>
              <oddFooter>&amp;Ca_x005F_x0041_b</oddFooter>
            </headerFooter>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_autofilter_with_escapes() {
        let mut worksheet = Worksheet::new();
        worksheet.autofilter(0, 0, 3, 1).unwrap();

        let filter = FilterCondition::new().add_list_filter("a_x0041_b");
        worksheet.filter_column(0, &filter).unwrap();

        let filter =
            FilterCondition::new().add_custom_filter(FilterCriteria::NotEqualTo, "a_x0041_b");
        worksheet.filter_column(1, &filter).unwrap();

        worksheet.write_auto_filter();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <autoFilter ref="A1:B4">
              <filterColumn colId="0">
                <filters>
                  <filter val="a_x005f_x0041_b"/>
                </filters>
              </filterColumn>
              <filterColumn colId="1">
                <customFilters>
                  <customFilter operator="notEqual" val="a_x005f_x0041_b"/>
                </customFilters>
              </filterColumn>
            </autoFilter>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_data_validation_with_escapes() {
        let mut worksheet = Worksheet::new();

        let data_validation = DataValidation::new()
            .allow_custom(r#"A1="a_x0041_b""#.into())
            .set_input_title("a_x0041_b")
            .unwrap();

        worksheet
            .add_data_validation(0, 0, 0, 0, &data_validation)
            .unwrap();

        worksheet.write_data_validations();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <dataValidations count="1">
              <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="a_x005f_x0041_b" sqref="A1">
                <formula1>A1="a_x005F_x0041_b"</formula1>
              </dataValidation>
            </dataValidations>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_hyperlink_with_escapes() {
        let mut worksheet = Worksheet::new();

        let url = Url::new("internal:a_x0041_b!A1").set_tip("a_x0041_b");
        worksheet.write_url(0, 0, &url).unwrap();

        worksheet.write_hyperlinks();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <hyperlinks>
              <hyperlink ref="A1" location="a_x005f_x0041_b!A1" tooltip="a_x005f_x0041_b" display="a_x005f_x0041_b!A1"/>
            </hyperlinks>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_sparkline_with_escapes() {
        let mut worksheet = Worksheet::new();

        let sparkline = Sparkline::new().set_range(("a_x0041_b", 0, 0, 0, 4));
        worksheet.add_sparkline(0, 5, &sparkline).unwrap();

        worksheet.write_sparkline_groups();

        let got = xmlwriter::cursor_to_str(&worksheet.writer);

        assert!(got.contains("<xm:f>a_x005F_x0041_b!A1:E1</xm:f>"), "{got}");
    }
    // Get the autofit column widths in column order.
    fn autofit_widths(worksheet: &Worksheet) -> Vec<(ColNum, u32)> {
        let mut widths: Vec<(ColNum, u32)> = worksheet
            .changed_cols
            .iter()
            .filter(|(_, col_options)| col_options.autofit)
            .map(|(col, col_options)| (*col, col_options.width))
            .collect();
        widths.sort_unstable();
        widths
    }

    #[test]
    fn test_autofit_empty_worksheet() {
        let mut worksheet = Worksheet::new();

        // Check for no panic and no column changes with an empty worksheet.
        worksheet.autofit();

        assert_eq!(autofit_widths(&worksheet), vec![]);
    }

    #[test]
    fn test_autofit_empty_constant_memory_worksheet() {
        let mut worksheet = Worksheet::new();
        worksheet.use_constant_memory = true;

        // Check for no panic and no column changes with an empty worksheet.
        worksheet.autofit();

        assert_eq!(autofit_widths(&worksheet), vec![]);
    }

    #[test]
    fn test_autofit_sparse_data() {
        let mut worksheet = Worksheet::new();

        // Write a diagonal of strings and some cells far from the others.
        for i in 0..20 {
            worksheet
                .write_string(i, i as ColNum, "x".repeat(i as usize + 1))
                .unwrap();
        }
        worksheet.write_number(100, 30, 12345).unwrap();
        worksheet.write_string(2, 16_000, "Far away").unwrap();
        worksheet.write_blank(3, 16_383, &Format::new()).unwrap();

        worksheet.autofit();

        // Check for the widths of the stored cells only.
        let mut expected: Vec<(ColNum, u32)> =
            (0..20).map(|col| (col, 14 + 7 * u32::from(col))).collect();
        expected.push((30, 42));
        expected.push((16_000, 61));

        assert_eq!(autofit_widths(&worksheet), expected);
    }

    #[test]
    fn test_autofit_max_row() {
        let mut worksheet = Worksheet::new();

        worksheet.write_string(0, 0, "Short").unwrap();
        worksheet
            .write_string(5, 0, "A much longer string")
            .unwrap();
        worksheet.write_string(6, 1, "Ignored").unwrap();

        worksheet.set_autofit_max_row(5);
        worksheet.autofit();

        // Check for the widths of rows before the max autofit row only.
        assert_eq!(autofit_widths(&worksheet), vec![(0, 40)]);
    }
}
