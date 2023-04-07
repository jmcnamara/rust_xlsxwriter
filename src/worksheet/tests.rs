// worksheet unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod worksheet_tests {

    use crate::test_functions::xml_to_vec;
    use crate::worksheet::SharedStringsTable;
    use crate::worksheet::{
        prepare_formula, FilterCondition, Format, HeaderImagePosition, Image, NaiveDate, NaiveTime,
        Worksheet, COL_MAX, ROW_MAX,
    };
    use crate::XlsxError;
    use pretty_assertions::assert_eq;
    use std::collections::HashMap;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet::default();
        let mut string_table = SharedStringsTable::new();

        worksheet.selected = true;

        worksheet.assemble_xml_file(&mut string_table);

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
        let worksheet = Worksheet::new();

        let strings = [
            ("", HeaderImagePosition::Left, false),
            ("&L&[Picture]", HeaderImagePosition::Left, true),
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
        ];

        for (string, position, exp) in strings {
            assert_eq!(exp, worksheet.verify_header_footer_image(string, &position));
        }
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
        let got = worksheet.process_pagebreaks(&[1, 1, 1, 1]).unwrap();
        assert_eq!(vec![1], got);

        // Test removing 0.
        let got = worksheet.process_pagebreaks(&[0, 1, 2, 3, 4]).unwrap();
        assert_eq!(vec![1, 2, 3, 4], got);

        // Test sort order.
        let got = worksheet.process_pagebreaks(&[1, 12, 2, 13, 3, 4]).unwrap();
        assert_eq!(vec![1, 2, 3, 4, 12, 13], got);

        // Exceed the number of allow breaks.
        let breaks = (1u32..=1024).collect::<Vec<u32>>();
        let result = worksheet.process_pagebreaks(&breaks);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));

        // Test row and column limits.
        let result = worksheet.set_page_breaks(&[ROW_MAX]);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.set_vertical_page_breaks(&[u32::from(COL_MAX)]);
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
    fn test_dynamic_function_escapes() {
        let formulas = vec![
            // Test simple escapes for formulas.
            ("=foo()", "foo()"),
            ("{foo()}", "foo()"),
            ("{=foo()}", "foo()"),
            // Dynamic functions.
            ("LET()", "_xlfn.LET()"),
            ("SEQUENCE(10)", "_xlfn.SEQUENCE(10)"),
            ("UNIQUES(A1:A10)", "UNIQUES(A1:A10)"),
            ("UUNIQUE(A1:A10)", "UUNIQUE(A1:A10)"),
            ("SINGLE(A1:A3)", "_xlfn.SINGLE(A1:A3)"),
            ("UNIQUE(A1:A10)", "_xlfn.UNIQUE(A1:A10)"),
            ("_xlfn.SEQUENCE(10)", "_xlfn.SEQUENCE(10)"),
            ("SORT(A1:A10)", "_xlfn._xlws.SORT(A1:A10)"),
            ("RANDARRAY(10,1)", "_xlfn.RANDARRAY(10,1)"),
            ("ANCHORARRAY(C1)", "_xlfn.ANCHORARRAY(C1)"),
            ("SORTBY(A1:A10,B1)", "_xlfn.SORTBY(A1:A10,B1)"),
            ("FILTER(A1:A10,1)", "_xlfn._xlws.FILTER(A1:A10,1)"),
            ("XMATCH(B1:B2,A1:A10)", "_xlfn.XMATCH(B1:B2,A1:A10)"),
            ("COUNTA(ANCHORARRAY(C1))", "COUNTA(_xlfn.ANCHORARRAY(C1))"),
            (
                "SEQUENCE(10)*SEQUENCE(10)",
                "_xlfn.SEQUENCE(10)*_xlfn.SEQUENCE(10)",
            ),
            (
                "XLOOKUP(\"India\",A22:A23,B22:B23)",
                "_xlfn.XLOOKUP(\"India\",A22:A23,B22:B23)",
            ),
            (
                "XLOOKUP(B1,A1:A10,ANCHORARRAY(D1))",
                "_xlfn.XLOOKUP(B1,A1:A10,_xlfn.ANCHORARRAY(D1))",
            ),
            (
                "LAMBDA(_xlpm.number, _xlpm.number + 1)(1)",
                "_xlfn.LAMBDA(_xlpm.number, _xlpm.number + 1)(1)",
            ),
            // Future functions.
            ("COT()", "_xlfn.COT()"),
            ("CSC()", "_xlfn.CSC()"),
            ("IFS()", "_xlfn.IFS()"),
            ("PHI()", "_xlfn.PHI()"),
            ("RRI()", "_xlfn.RRI()"),
            ("SEC()", "_xlfn.SEC()"),
            ("XOR()", "_xlfn.XOR()"),
            ("ACOT()", "_xlfn.ACOT()"),
            ("BASE()", "_xlfn.BASE()"),
            ("COTH()", "_xlfn.COTH()"),
            ("CSCH()", "_xlfn.CSCH()"),
            ("DAYS()", "_xlfn.DAYS()"),
            ("IFNA()", "_xlfn.IFNA()"),
            ("SECH()", "_xlfn.SECH()"),
            ("ACOTH()", "_xlfn.ACOTH()"),
            ("BITOR()", "_xlfn.BITOR()"),
            ("F.INV()", "_xlfn.F.INV()"),
            ("GAMMA()", "_xlfn.GAMMA()"),
            ("GAUSS()", "_xlfn.GAUSS()"),
            ("IMCOT()", "_xlfn.IMCOT()"),
            ("IMCSC()", "_xlfn.IMCSC()"),
            ("IMSEC()", "_xlfn.IMSEC()"),
            ("IMTAN()", "_xlfn.IMTAN()"),
            ("MUNIT()", "_xlfn.MUNIT()"),
            ("SHEET()", "_xlfn.SHEET()"),
            ("T.INV()", "_xlfn.T.INV()"),
            ("VAR.P()", "_xlfn.VAR.P()"),
            ("VAR.S()", "_xlfn.VAR.S()"),
            ("ARABIC()", "_xlfn.ARABIC()"),
            ("BITAND()", "_xlfn.BITAND()"),
            ("BITXOR()", "_xlfn.BITXOR()"),
            ("CONCAT()", "_xlfn.CONCAT()"),
            ("F.DIST()", "_xlfn.F.DIST()"),
            ("F.TEST()", "_xlfn.F.TEST()"),
            ("IMCOSH()", "_xlfn.IMCOSH()"),
            ("IMCSCH()", "_xlfn.IMCSCH()"),
            ("IMSECH()", "_xlfn.IMSECH()"),
            ("IMSINH()", "_xlfn.IMSINH()"),
            ("MAXIFS()", "_xlfn.MAXIFS()"),
            ("MINIFS()", "_xlfn.MINIFS()"),
            ("SHEETS()", "_xlfn.SHEETS()"),
            ("SKEW.P()", "_xlfn.SKEW.P()"),
            ("SWITCH()", "_xlfn.SWITCH()"),
            ("T.DIST()", "_xlfn.T.DIST()"),
            ("T.TEST()", "_xlfn.T.TEST()"),
            ("Z.TEST()", "_xlfn.Z.TEST()"),
            ("COMBINA()", "_xlfn.COMBINA()"),
            ("DECIMAL()", "_xlfn.DECIMAL()"),
            ("RANK.EQ()", "_xlfn.RANK.EQ()"),
            ("STDEV.P()", "_xlfn.STDEV.P()"),
            ("STDEV.S()", "_xlfn.STDEV.S()"),
            ("UNICHAR()", "_xlfn.UNICHAR()"),
            ("UNICODE()", "_xlfn.UNICODE()"),
            ("BETA.INV()", "_xlfn.BETA.INV()"),
            ("F.INV.RT()", "_xlfn.F.INV.RT()"),
            ("ISO.CEILING()", "ISO.CEILING()"),
            ("NORM.INV()", "_xlfn.NORM.INV()"),
            ("RANK.AVG()", "_xlfn.RANK.AVG()"),
            ("T.INV.2T()", "_xlfn.T.INV.2T()"),
            ("TEXTJOIN()", "_xlfn.TEXTJOIN()"),
            ("AGGREGATE()", "_xlfn.AGGREGATE()"),
            ("BETA.DIST()", "_xlfn.BETA.DIST()"),
            ("BINOM.INV()", "_xlfn.BINOM.INV()"),
            ("BITLSHIFT()", "_xlfn.BITLSHIFT()"),
            ("BITRSHIFT()", "_xlfn.BITRSHIFT()"),
            ("CHISQ.INV()", "_xlfn.CHISQ.INV()"),
            ("ECMA.CEILING()", "ECMA.CEILING()"),
            ("F.DIST.RT()", "_xlfn.F.DIST.RT()"),
            ("FILTERXML()", "_xlfn.FILTERXML()"),
            ("GAMMA.INV()", "_xlfn.GAMMA.INV()"),
            ("ISFORMULA()", "_xlfn.ISFORMULA()"),
            ("MODE.MULT()", "_xlfn.MODE.MULT()"),
            ("MODE.SNGL()", "_xlfn.MODE.SNGL()"),
            ("NORM.DIST()", "_xlfn.NORM.DIST()"),
            ("PDURATION()", "_xlfn.PDURATION()"),
            ("T.DIST.2T()", "_xlfn.T.DIST.2T()"),
            ("T.DIST.RT()", "_xlfn.T.DIST.RT()"),
            ("WORKDAY.INTL()", "WORKDAY.INTL()"),
            ("BINOM.DIST()", "_xlfn.BINOM.DIST()"),
            ("CHISQ.DIST()", "_xlfn.CHISQ.DIST()"),
            ("CHISQ.TEST()", "_xlfn.CHISQ.TEST()"),
            ("EXPON.DIST()", "_xlfn.EXPON.DIST()"),
            ("FLOOR.MATH()", "_xlfn.FLOOR.MATH()"),
            ("GAMMA.DIST()", "_xlfn.GAMMA.DIST()"),
            ("ISOWEEKNUM()", "_xlfn.ISOWEEKNUM()"),
            ("NORM.S.INV()", "_xlfn.NORM.S.INV()"),
            ("WEBSERVICE()", "_xlfn.WEBSERVICE()"),
            ("ERF.PRECISE()", "_xlfn.ERF.PRECISE()"),
            ("FORMULATEXT()", "_xlfn.FORMULATEXT()"),
            ("LOGNORM.INV()", "_xlfn.LOGNORM.INV()"),
            ("NORM.S.DIST()", "_xlfn.NORM.S.DIST()"),
            ("NUMBERVALUE()", "_xlfn.NUMBERVALUE()"),
            ("QUERYSTRING()", "_xlfn.QUERYSTRING()"),
            ("CEILING.MATH()", "_xlfn.CEILING.MATH()"),
            ("CHISQ.INV.RT()", "_xlfn.CHISQ.INV.RT()"),
            ("CONFIDENCE.T()", "_xlfn.CONFIDENCE.T()"),
            ("COVARIANCE.P()", "_xlfn.COVARIANCE.P()"),
            ("COVARIANCE.S()", "_xlfn.COVARIANCE.S()"),
            ("ERFC.PRECISE()", "_xlfn.ERFC.PRECISE()"),
            ("FORECAST.ETS()", "_xlfn.FORECAST.ETS()"),
            ("HYPGEOM.DIST()", "_xlfn.HYPGEOM.DIST()"),
            ("LOGNORM.DIST()", "_xlfn.LOGNORM.DIST()"),
            ("PERMUTATIONA()", "_xlfn.PERMUTATIONA()"),
            ("POISSON.DIST()", "_xlfn.POISSON.DIST()"),
            ("QUARTILE.EXC()", "_xlfn.QUARTILE.EXC()"),
            ("QUARTILE.INC()", "_xlfn.QUARTILE.INC()"),
            ("WEIBULL.DIST()", "_xlfn.WEIBULL.DIST()"),
            ("CHISQ.DIST.RT()", "_xlfn.CHISQ.DIST.RT()"),
            ("FLOOR.PRECISE()", "_xlfn.FLOOR.PRECISE()"),
            ("NEGBINOM.DIST()", "_xlfn.NEGBINOM.DIST()"),
            ("NETWORKDAYS.INTL()", "NETWORKDAYS.INTL()"),
            ("PERCENTILE.EXC()", "_xlfn.PERCENTILE.EXC()"),
            ("PERCENTILE.INC()", "_xlfn.PERCENTILE.INC()"),
            ("CEILING.PRECISE()", "_xlfn.CEILING.PRECISE()"),
            ("CONFIDENCE.NORM()", "_xlfn.CONFIDENCE.NORM()"),
            ("FORECAST.LINEAR()", "_xlfn.FORECAST.LINEAR()"),
            ("GAMMALN.PRECISE()", "_xlfn.GAMMALN.PRECISE()"),
            ("PERCENTRANK.EXC()", "_xlfn.PERCENTRANK.EXC()"),
            ("PERCENTRANK.INC()", "_xlfn.PERCENTRANK.INC()"),
            ("BINOM.DIST.RANGE()", "_xlfn.BINOM.DIST.RANGE()"),
            ("FORECAST.ETS.STAT()", "_xlfn.FORECAST.ETS.STAT()"),
            ("FORECAST.ETS.CONFINT()", "_xlfn.FORECAST.ETS.CONFINT()"),
            (
                "FORECAST.ETS.SEASONALITY()",
                "_xlfn.FORECAST.ETS.SEASONALITY()",
            ),
            (
                "Z.TEST(Z.TEST(Z.TEST()))",
                "_xlfn.Z.TEST(_xlfn.Z.TEST(_xlfn.Z.TEST()))",
            ),
        ];

        for test_data in &formulas {
            let mut formula = test_data.0.to_string();
            let expected = test_data.1;

            formula = prepare_formula(&formula, true);

            assert_eq!(formula, expected);
        }
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
        assert!(matches!(result, Err(XlsxError::SheetnameCannotBeBlank)));

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

    #[test]
    fn dates_and_times() {
        let mut worksheet = Worksheet::new();

        // Test date and time
        #[allow(clippy::excessive_precision)]
        let datetimes = vec![
            (1899, 12, 31, 0, 0, 0, 0, 0.0),
            (1982, 8, 25, 0, 15, 20, 213, 30_188.010_650_613_425),
            (2065, 4, 19, 0, 16, 48, 290, 60_376.011_670_023_145),
            (2147, 12, 15, 0, 55, 25, 446, 90_565.038_488_958_337),
            (2230, 8, 10, 1, 2, 46, 891, 120_753.043_598_275_46),
            (2313, 4, 6, 1, 4, 15, 597, 150_942.044_624_965_29),
            (2395, 11, 30, 1, 9, 40, 889, 181_130.048_389_918_99),
            (2478, 7, 25, 1, 11, 32, 560, 211_318.049_682_407_41),
            (2561, 3, 21, 1, 30, 19, 169, 241_507.062_721_863_42),
            (2643, 11, 15, 1, 48, 25, 580, 271_695.075_296_064_84),
            (2726, 7, 12, 2, 3, 31, 919, 301_884.085_786_099_55),
            (2809, 3, 6, 2, 11, 11, 986, 332_072.091_110_949_06),
            (2891, 10, 31, 2, 24, 37, 95, 362_261.100_429_340_27),
            (2974, 6, 26, 2, 35, 7, 220, 392_449.107_722_453_71),
            (3057, 2, 19, 2, 45, 12, 109, 422_637.114_723_483_8),
            (3139, 10, 17, 3, 6, 39, 990, 452_826.129_629_513_89),
            (3222, 6, 11, 3, 8, 8, 251, 483_014.130_651_053_22),
            (3305, 2, 5, 3, 19, 12, 576, 513_203.138_34),
            (3387, 10, 1, 3, 29, 42, 574, 543_391.145_631_643_48),
            (3470, 5, 27, 3, 37, 30, 813, 573_579.151_051_076_36),
            (3553, 1, 21, 4, 14, 38, 231, 603_768.176_831_377_32),
            (3635, 9, 16, 4, 16, 28, 559, 633_956.178_108_321_74),
            (3718, 5, 13, 4, 17, 58, 222, 664_145.179_146_087_96),
            (3801, 1, 6, 4, 21, 41, 794, 694_333.181_733_726_87),
            (3883, 9, 2, 4, 56, 35, 792, 724_522.205_969_814_79),
            (3966, 4, 28, 5, 25, 14, 885, 754_710.225_866_724_5),
            (4048, 12, 21, 5, 26, 5, 724, 784_898.226_455_138_88),
            (4131, 8, 18, 5, 46, 44, 68, 815_087.240_787_824_03),
            (4214, 4, 13, 5, 48, 1, 141, 845_275.241_679_872_74),
            (4296, 12, 7, 5, 53, 52, 315, 875_464.245_744_386_57),
            (4379, 8, 3, 6, 14, 48, 580, 905_652.260_284_490_77),
            (4462, 3, 28, 6, 46, 15, 738, 935_840.282_126_597_25),
            (4544, 11, 22, 7, 31, 20, 407, 966_029.313_430_636_54),
            (4627, 7, 19, 7, 58, 33, 754, 996_217.332_335_115_76),
            (4710, 3, 15, 8, 7, 43, 130, 1_026_406.338_693_634_3),
            (4792, 11, 7, 8, 29, 11, 91, 1_056_594.353_600_590_3),
            (4875, 7, 4, 9, 8, 15, 328, 1_086_783.380_732_962_9),
            (4958, 2, 27, 9, 30, 41, 781, 1_116_971.396_316_909_7),
            (5040, 10, 23, 9, 34, 4, 462, 1_147_159.398_662_754_6),
            (5123, 6, 20, 9, 37, 23, 945, 1_177_348.400_971_585_7),
            (5206, 2, 12, 9, 37, 56, 655, 1_207_536.401_350_173_6),
            (5288, 10, 8, 9, 45, 12, 230, 1_237_725.406_391_551),
            (5371, 6, 4, 9, 54, 14, 782, 1_267_913.412_671_088),
            (5454, 1, 28, 9, 54, 22, 108, 1_298_101.412_755_879_6),
            (5536, 9, 24, 10, 1, 36, 151, 1_328_290.417_779_525_5),
            (5619, 5, 20, 12, 9, 48, 602, 1_358_478.506_812_523_1),
            (5702, 1, 14, 12, 34, 8, 549, 1_388_667.523_710_057_8),
            (5784, 9, 8, 12, 56, 6, 495, 1_418_855.538_964_062_5),
            (5867, 5, 6, 12, 58, 58, 217, 1_449_044.540_951_585_6),
            (5949, 12, 30, 12, 59, 54, 263, 1_479_232.541_600_266_2),
            (6032, 8, 24, 13, 34, 41, 331, 1_509_420.565_756_145_9),
            (6115, 4, 21, 13, 58, 28, 601, 1_539_609.582_275_474_4),
            (6197, 12, 14, 14, 2, 16, 899, 1_569_797.584_917_812_6),
            (6280, 8, 10, 14, 36, 17, 444, 1_599_986.608_535_231_6),
            (6363, 4, 6, 14, 37, 57, 451, 1_630_174.609_692_72),
            (6445, 11, 30, 14, 57, 42, 757, 1_660_363.623_411_539_2),
            (6528, 7, 26, 15, 10, 48, 307, 1_690_551.632_503_553_3),
            (6611, 3, 22, 15, 14, 39, 890, 1_720_739.635_183_912),
            (6693, 11, 15, 15, 19, 47, 988, 1_750_928.638_749_861_2),
            (6776, 7, 11, 16, 4, 24, 344, 1_781_116.669_726_203_7),
            (6859, 3, 7, 16, 22, 23, 952, 1_811_305.682_221_666_7),
            (6941, 10, 31, 16, 29, 55, 999, 1_841_493.687_453_692_1),
            (7024, 6, 26, 16, 58, 20, 259, 1_871_681.707_178_923_5),
            (7107, 2, 21, 17, 4, 2, 415, 1_901_870.711_139_062_4),
            (7189, 10, 16, 17, 18, 29, 630, 1_932_058.721_176_273_2),
            (7272, 6, 11, 17, 47, 21, 323, 1_962_247.741_219_016_3),
            (7355, 2, 5, 17, 53, 29, 866, 1_992_435.745_484_560_3),
            (7437, 10, 2, 17, 53, 41, 76, 2_022_624.745_614_305_6),
            (7520, 5, 28, 17, 55, 6, 44, 2_052_812.746_597_731_5),
            (7603, 1, 21, 18, 14, 49, 151, 2_083_000.760_291_099_5),
            (7685, 9, 16, 18, 17, 45, 738, 2_113_189.762_334_930_7),
            (7768, 5, 12, 18, 29, 59, 700, 2_143_377.770_829_861_1),
            (7851, 1, 7, 18, 33, 21, 233, 2_173_566.773_162_419),
            (7933, 9, 2, 19, 14, 24, 673, 2_203_754.801_674_455_9),
            (8016, 4, 27, 19, 17, 12, 816, 2_233_942.803_620_555_4),
            (8098, 12, 22, 19, 23, 36, 418, 2_264_131.808_060_393_7),
            (8181, 8, 17, 19, 46, 25, 908, 2_294_319.823_910_972_1),
            (8264, 4, 13, 20, 7, 47, 314, 2_324_508.838_742_060_1),
            (8346, 12, 8, 20, 31, 37, 603, 2_354_696.855_296_331),
            (8429, 8, 3, 20, 39, 57, 770, 2_384_885.861_085_300_8),
            (8512, 3, 29, 20, 50, 17, 67, 2_415_073.868_253_090_4),
            (8594, 11, 22, 21, 2, 57, 827, 2_445_261.877_058_182_8),
            (8677, 7, 19, 21, 23, 5, 519, 2_475_450.891_036_099_8),
            (8760, 3, 14, 21, 34, 49, 572, 2_505_638.899_184_861_2),
            (8842, 11, 8, 21, 39, 5, 944, 2_535_827.902_152_129_4),
            (8925, 7, 4, 21, 39, 18, 426, 2_566_015.902_296_597_1),
            (9008, 2, 28, 21, 46, 7, 769, 2_596_203.907_034_363_6),
            (9090, 10, 24, 21, 57, 55, 662, 2_626_392.915_227_569_6),
            (9173, 6, 19, 22, 19, 11, 732, 2_656_580.929_996_897_9),
            (9256, 2, 13, 22, 23, 51, 376, 2_686_769.933_233_518_6),
            (9338, 10, 9, 22, 27, 58, 771, 2_716_957.936_096_886_6),
            (9421, 6, 5, 22, 43, 30, 392, 2_747_146.946_879_536_8),
            (9504, 1, 30, 22, 48, 25, 834, 2_777_334.950_299_004_6),
            (9586, 9, 24, 22, 53, 51, 727, 2_807_522.954_070_914_5),
            (9669, 5, 20, 23, 12, 56, 536, 2_837_711.967_321_018_7),
            (9752, 1, 14, 23, 15, 54, 109, 2_867_899.969_376_261_3),
            (9834, 9, 10, 23, 17, 12, 632, 2_898_088.970_285_092_5),
            (9999, 12, 31, 23, 59, 59, 0, 2_958_465.999_988_426),
        ];

        for test_data in datetimes {
            let (year, month, day, hour, min, seconds, millis, expected) = test_data;
            let datetime = NaiveDate::from_ymd_opt(year, month, day)
                .unwrap()
                .and_hms_milli_opt(hour, min, seconds, millis)
                .unwrap();
            assert_eq!(expected, worksheet.datetime_to_excel(&datetime));
        }
    }

    #[test]
    fn dates_only() {
        let mut worksheet = Worksheet::new();

        // Test date only.
        let dates = vec![
            (1899, 12, 31, 0.0),
            (1900, 1, 1, 1.0),
            (1900, 2, 27, 58.0),
            (1900, 2, 28, 59.0),
            (1900, 3, 1, 61.0),
            (1900, 3, 2, 62.0),
            (1900, 3, 11, 71.0),
            (1900, 4, 8, 99.0),
            (1900, 9, 12, 256.0),
            (1901, 5, 3, 489.0),
            (1901, 10, 13, 652.0),
            (1902, 2, 15, 777.0),
            (1902, 6, 6, 888.0),
            (1902, 9, 25, 999.0),
            (1902, 9, 27, 1001.0),
            (1903, 4, 26, 1212.0),
            (1903, 8, 5, 1313.0),
            (1903, 12, 31, 1461.0),
            (1904, 1, 1, 1462.0),
            (1904, 2, 28, 1520.0),
            (1904, 2, 29, 1521.0),
            (1904, 3, 1, 1522.0),
            (1907, 2, 27, 2615.0),
            (1907, 2, 28, 2616.0),
            (1907, 3, 1, 2617.0),
            (1907, 3, 2, 2618.0),
            (1907, 3, 3, 2619.0),
            (1907, 3, 4, 2620.0),
            (1907, 3, 5, 2621.0),
            (1907, 3, 6, 2622.0),
            (1999, 1, 1, 36161.0),
            (1999, 1, 31, 36191.0),
            (1999, 2, 1, 36192.0),
            (1999, 2, 28, 36219.0),
            (1999, 3, 1, 36220.0),
            (1999, 3, 31, 36250.0),
            (1999, 4, 1, 36251.0),
            (1999, 4, 30, 36280.0),
            (1999, 5, 1, 36281.0),
            (1999, 5, 31, 36311.0),
            (1999, 6, 1, 36312.0),
            (1999, 6, 30, 36341.0),
            (1999, 7, 1, 36342.0),
            (1999, 7, 31, 36372.0),
            (1999, 8, 1, 36373.0),
            (1999, 8, 31, 36403.0),
            (1999, 9, 1, 36404.0),
            (1999, 9, 30, 36433.0),
            (1999, 10, 1, 36434.0),
            (1999, 10, 31, 36464.0),
            (1999, 11, 1, 36465.0),
            (1999, 11, 30, 36494.0),
            (1999, 12, 1, 36495.0),
            (1999, 12, 31, 36525.0),
            (2000, 1, 1, 36526.0),
            (2000, 1, 31, 36556.0),
            (2000, 2, 1, 36557.0),
            (2000, 2, 29, 36585.0),
            (2000, 3, 1, 36586.0),
            (2000, 3, 31, 36616.0),
            (2000, 4, 1, 36617.0),
            (2000, 4, 30, 36646.0),
            (2000, 5, 1, 36647.0),
            (2000, 5, 31, 36677.0),
            (2000, 6, 1, 36678.0),
            (2000, 6, 30, 36707.0),
            (2000, 7, 1, 36708.0),
            (2000, 7, 31, 36738.0),
            (2000, 8, 1, 36739.0),
            (2000, 8, 31, 36769.0),
            (2000, 9, 1, 36770.0),
            (2000, 9, 30, 36799.0),
            (2000, 10, 1, 36800.0),
            (2000, 10, 31, 36830.0),
            (2000, 11, 1, 36831.0),
            (2000, 11, 30, 36860.0),
            (2000, 12, 1, 36861.0),
            (2000, 12, 31, 36891.0),
            (2001, 1, 1, 36892.0),
            (2001, 1, 31, 36922.0),
            (2001, 2, 1, 36923.0),
            (2001, 2, 28, 36950.0),
            (2001, 3, 1, 36951.0),
            (2001, 3, 31, 36981.0),
            (2001, 4, 1, 36982.0),
            (2001, 4, 30, 37011.0),
            (2001, 5, 1, 37012.0),
            (2001, 5, 31, 37042.0),
            (2001, 6, 1, 37043.0),
            (2001, 6, 30, 37072.0),
            (2001, 7, 1, 37073.0),
            (2001, 7, 31, 37103.0),
            (2001, 8, 1, 37104.0),
            (2001, 8, 31, 37134.0),
            (2001, 9, 1, 37135.0),
            (2001, 9, 30, 37164.0),
            (2001, 10, 1, 37165.0),
            (2001, 10, 31, 37195.0),
            (2001, 11, 1, 37196.0),
            (2001, 11, 30, 37225.0),
            (2001, 12, 1, 37226.0),
            (2001, 12, 31, 37256.0),
            (2400, 1, 1, 182_623.0),
            (2400, 1, 31, 182_653.0),
            (2400, 2, 1, 182_654.0),
            (2400, 2, 29, 182_682.0),
            (2400, 3, 1, 182_683.0),
            (2400, 3, 31, 182_713.0),
            (2400, 4, 1, 182_714.0),
            (2400, 4, 30, 182_743.0),
            (2400, 5, 1, 182_744.0),
            (2400, 5, 31, 182_774.0),
            (2400, 6, 1, 182_775.0),
            (2400, 6, 30, 182_804.0),
            (2400, 7, 1, 182_805.0),
            (2400, 7, 31, 182_835.0),
            (2400, 8, 1, 182_836.0),
            (2400, 8, 31, 182_866.0),
            (2400, 9, 1, 182_867.0),
            (2400, 9, 30, 182_896.0),
            (2400, 10, 1, 182_897.0),
            (2400, 10, 31, 182_927.0),
            (2400, 11, 1, 182_928.0),
            (2400, 11, 30, 182_957.0),
            (2400, 12, 1, 182_958.0),
            (2400, 12, 31, 182_988.0),
            (4000, 1, 1, 767_011.0),
            (4000, 1, 31, 767_041.0),
            (4000, 2, 1, 767_042.0),
            (4000, 2, 29, 767_070.0),
            (4000, 3, 1, 767_071.0),
            (4000, 3, 31, 767_101.0),
            (4000, 4, 1, 767_102.0),
            (4000, 4, 30, 767_131.0),
            (4000, 5, 1, 767_132.0),
            (4000, 5, 31, 767_162.0),
            (4000, 6, 1, 767_163.0),
            (4000, 6, 30, 767_192.0),
            (4000, 7, 1, 767_193.0),
            (4000, 7, 31, 767_223.0),
            (4000, 8, 1, 767_224.0),
            (4000, 8, 31, 767_254.0),
            (4000, 9, 1, 767_255.0),
            (4000, 9, 30, 767_284.0),
            (4000, 10, 1, 767_285.0),
            (4000, 10, 31, 767_315.0),
            (4000, 11, 1, 767_316.0),
            (4000, 11, 30, 767_345.0),
            (4000, 12, 1, 767_346.0),
            (4000, 12, 31, 767_376.0),
            (4321, 1, 1, 884_254.0),
            (4321, 1, 31, 884_284.0),
            (4321, 2, 1, 884_285.0),
            (4321, 2, 28, 884_312.0),
            (4321, 3, 1, 884_313.0),
            (4321, 3, 31, 884_343.0),
            (4321, 4, 1, 884_344.0),
            (4321, 4, 30, 884_373.0),
            (4321, 5, 1, 884_374.0),
            (4321, 5, 31, 884_404.0),
            (4321, 6, 1, 884_405.0),
            (4321, 6, 30, 884_434.0),
            (4321, 7, 1, 884_435.0),
            (4321, 7, 31, 884_465.0),
            (4321, 8, 1, 884_466.0),
            (4321, 8, 31, 884_496.0),
            (4321, 9, 1, 884_497.0),
            (4321, 9, 30, 884_526.0),
            (4321, 10, 1, 884_527.0),
            (4321, 10, 31, 884_557.0),
            (4321, 11, 1, 884_558.0),
            (4321, 11, 30, 884_587.0),
            (4321, 12, 1, 884_588.0),
            (4321, 12, 31, 884_618.0),
            (9999, 1, 1, 2_958_101.0),
            (9999, 1, 31, 2_958_131.0),
            (9999, 2, 1, 2_958_132.0),
            (9999, 2, 28, 2_958_159.0),
            (9999, 3, 1, 2_958_160.0),
            (9999, 3, 31, 2_958_190.0),
            (9999, 4, 1, 2_958_191.0),
            (9999, 4, 30, 2_958_220.0),
            (9999, 5, 1, 2_958_221.0),
            (9999, 5, 31, 2_958_251.0),
            (9999, 6, 1, 2_958_252.0),
            (9999, 6, 30, 2_958_281.0),
            (9999, 7, 1, 2_958_282.0),
            (9999, 7, 31, 2_958_312.0),
            (9999, 8, 1, 2_958_313.0),
            (9999, 8, 31, 2_958_343.0),
            (9999, 9, 1, 2_958_344.0),
            (9999, 9, 30, 2_958_373.0),
            (9999, 10, 1, 2_958_374.0),
            (9999, 10, 31, 2_958_404.0),
            (9999, 11, 1, 2_958_405.0),
            (9999, 11, 30, 2_958_434.0),
            (9999, 12, 1, 2_958_435.0),
            (9999, 12, 31, 2_958_465.0),
        ];

        for test_data in dates {
            let (year, month, day, expected) = test_data;
            let datetime = NaiveDate::from_ymd_opt(year, month, day).unwrap();
            assert_eq!(expected, worksheet.date_to_excel(&datetime));
        }
    }

    #[test]
    fn times_only() {
        let mut worksheet = Worksheet::new();

        // Test time only.
        #[allow(clippy::excessive_precision)]
        let times = vec![
            (0, 0, 0, 0, 0.0),
            (0, 15, 20, 213, 1.065_061_342_592_592_4E-2),
            (0, 16, 48, 290, 1.167_002_314_814_814_8E-2),
            (0, 55, 25, 446, 3.848_895_833_333_333_7E-2),
            (1, 2, 46, 891, 4.359_827_546_296_296_5E-2),
            (1, 4, 15, 597, 4.462_496_527_777_778_2E-2),
            (1, 9, 40, 889, 4.838_991_898_148_148_3E-2),
            (1, 11, 32, 560, 4.968_240_740_740_740_4E-2),
            (1, 30, 19, 169, 6.272_186_342_592_593_6E-2),
            (1, 48, 25, 580, 7.529_606_481_481_480_9E-2),
            (2, 3, 31, 919, 8.578_609_953_703_703_1E-2),
            (2, 11, 11, 986, 9.111_094_907_407_407_7E-2),
            (2, 24, 37, 95, 0.100_429_340_277_777_78),
            (2, 35, 7, 220, 0.107_722_453_703_703_7),
            (2, 45, 12, 109, 0.114_723_483_796_296_31),
            (3, 6, 39, 990, 0.129_629_513_888_888_88),
            (3, 8, 8, 251, 0.130_651_053_240_740_75),
            (3, 19, 12, 576, 0.138_339_999_999_999_99),
            (3, 29, 42, 574, 0.145_631_643_518_518_51),
            (3, 37, 30, 813, 0.151_051_076_388_888_9),
            (4, 14, 38, 231, 0.176_831_377_314_814_8),
            (4, 16, 28, 559, 0.178_108_321_759_259_25),
            (4, 17, 58, 222, 0.179_146_087_962_962_97),
            (4, 21, 41, 794, 0.181_733_726_851_851_85),
            (4, 56, 35, 792, 0.205_969_814_814_814_8),
            (5, 25, 14, 885, 0.225_866_724_537_037_04),
            (5, 26, 5, 724, 0.226_455_138_888_888_91),
            (5, 46, 44, 68, 0.240_787_824_074_074_06),
            (5, 48, 1, 141, 0.241_679_872_685_185_2),
            (5, 53, 52, 315, 0.245_744_386_574_074_08),
            (6, 14, 48, 580, 0.260_284_490_740_740_73),
            (6, 46, 15, 738, 0.282_126_597_222_222_22),
            (7, 31, 20, 407, 0.313_430_636_574_074_05),
            (7, 58, 33, 754, 0.332_335_115_740_740_76),
            (8, 7, 43, 130, 0.338_693_634_259_259_25),
            (8, 29, 11, 91, 0.353_600_590_277_777_74),
            (9, 8, 15, 328, 0.380_732_962_962_963),
            (9, 30, 41, 781, 0.396_316_909_722_222_28),
            (9, 34, 4, 462, 0.398_662_754_629_629_58),
            (9, 37, 23, 945, 0.400_971_585_648_148_17),
            (9, 37, 56, 655, 0.401_350_173_611_111_14),
            (9, 45, 12, 230, 0.406_391_550_925_925_94),
            (9, 54, 14, 782, 0.412_671_087_962_962_98),
            (9, 54, 22, 108, 0.412_755_879_629_629_62),
            (10, 1, 36, 151, 0.417_779_525_462_962_99),
            (12, 9, 48, 602, 0.506_812_523_148_148_18),
            (12, 34, 8, 549, 0.523_710_057_870_370_39),
            (12, 56, 6, 495, 0.538_964_062_499_999_95),
            (12, 58, 58, 217, 0.540_951_585_648_148_16),
            (12, 59, 54, 263, 0.541_600_266_203_703_72),
            (13, 34, 41, 331, 0.565_756_145_833_333_33),
            (13, 58, 28, 601, 0.582_275_474_537_036_99),
            (14, 2, 16, 899, 0.584_917_812_499_999_97),
            (14, 36, 17, 444, 0.608_535_231_481_481_48),
            (14, 37, 57, 451, 0.609_692_719_907_407_48),
            (14, 57, 42, 757, 0.623_411_539_351_851_9),
            (15, 10, 48, 307, 0.632_503_553_240_740_7),
            (15, 14, 39, 890, 0.635_183_912_037_037_06),
            (15, 19, 47, 988, 0.638_749_861_111_111_09),
            (16, 4, 24, 344, 0.669_726_203_703_703_62),
            (16, 22, 23, 952, 0.682_221_666_666_666_62),
            (16, 29, 55, 999, 0.687_453_692_129_629_7),
            (16, 58, 20, 259, 0.707_178_923_611_111_12),
            (17, 4, 2, 415, 0.711_139_062_500_000_03),
            (17, 18, 29, 630, 0.721_176_273_148_148_25),
            (17, 47, 21, 323, 0.741_219_016_203_703_67),
            (17, 53, 29, 866, 0.745_484_560_185_185_16),
            (17, 53, 41, 76, 0.745_614_305_555_555_63),
            (17, 55, 6, 44, 0.746_597_731_481_481_45),
            (18, 14, 49, 151, 0.760_291_099_537_037),
            (18, 17, 45, 738, 0.762_334_930_555_555_46),
            (18, 29, 59, 700, 0.770_829_861_111_111_18),
            (18, 33, 21, 233, 0.773_162_418_981_481_53),
            (19, 14, 24, 673, 0.801_674_456_018_518_61),
            (19, 17, 12, 816, 0.803_620_555_555_555_45),
            (19, 23, 36, 418, 0.808_060_393_518_518_55),
            (19, 46, 25, 908, 0.823_910_972_222_222_32),
            (20, 7, 47, 314, 0.838_742_060_185_185_16),
            (20, 31, 37, 603, 0.855_296_331_018_518_54),
            (20, 39, 57, 770, 0.861_085_300_925_925_94),
            (20, 50, 17, 67, 0.868_253_090_277_777_75),
            (21, 2, 57, 827, 0.877_058_182_870_370_41),
            (21, 23, 5, 519, 0.891_036_099_537_037),
            (21, 34, 49, 572, 0.899_184_861_111_111_18),
            (21, 39, 5, 944, 0.902_152_129_629_629_65),
            (21, 39, 18, 426, 0.902_296_597_222_222_22),
            (21, 46, 7, 769, 0.907_034_363_425_926_03),
            (21, 57, 55, 662, 0.915_227_569_444_444_39),
            (22, 19, 11, 732, 0.929_996_898_148_148_23),
            (22, 23, 51, 376, 0.933_233_518_518_518_43),
            (22, 27, 58, 771, 0.936_096_886_574_074_08),
            (22, 43, 30, 392, 0.946_879_537_037_037_09),
            (22, 48, 25, 834, 0.950_299_004_629_629_68),
            (22, 53, 51, 727, 0.954_070_914_351_851_87),
            (23, 12, 56, 536, 0.967_321_018_518_518_48),
            (23, 15, 54, 109, 0.969_376_261_574_074_08),
            (23, 17, 12, 632, 0.970_285_092_592_592_66),
            (23, 59, 59, 999, 0.999_999_988_425_925_86),
        ];

        for test_data in times {
            let (hour, min, seconds, millis, expected) = test_data;
            let datetime = NaiveTime::from_hms_milli_opt(hour, min, seconds, millis).unwrap();
            let mut diff = worksheet.time_to_excel(&datetime) - expected;
            diff = diff.abs();
            assert!(diff < 0.000_000_000_01);
        }
    }
}
