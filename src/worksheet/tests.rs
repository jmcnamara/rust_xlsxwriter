// worksheet unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod tests {

    use crate::worksheet::SharedStringsTable;
    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
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
        let got = xml_to_vec(&got);

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

        assert_eq!(got, expected);
    }

    #[test]
    fn verify_header_footer_images() {
        let worksheet = Worksheet::new();

        let strings = [
            ("", XlsxImagePosition::Left, false),
            ("&L&[Picture]", XlsxImagePosition::Left, true),
            ("&R&[Picture]", XlsxImagePosition::Right, true),
            ("&C&[Picture]", XlsxImagePosition::Center, true),
            ("&R&[Picture]", XlsxImagePosition::Left, false),
            ("&L&[Picture]&C&[Picture]", XlsxImagePosition::Left, true),
            ("&L&[Picture]&C&[Picture]", XlsxImagePosition::Center, true),
            ("&L&[Picture]&C&[Picture]", XlsxImagePosition::Right, false),
        ];

        for (string, position, exp) in strings {
            assert_eq!(exp, worksheet.verify_header_footer_image(string, &position));
        }
    }

    #[test]
    fn row_matches_list_filter_blanks() {
        let mut worksheet = Worksheet::new();
        let bold = Format::new().set_bold();

        worksheet.write_string_only(0, 0, "Header").unwrap();
        worksheet.write_string_only(1, 0, "").unwrap();
        worksheet.write_string_only(2, 0, " ").unwrap();
        worksheet.write_string_only(3, 0, "  ").unwrap();
        worksheet.write_string(4, 0, "", &bold).unwrap();

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
        worksheet.write_string_only(0, 0, "Header").unwrap();
        worksheet.write_string_only(1, 0, "South").unwrap();
        worksheet.write_string_only(2, 0, "south").unwrap();
        worksheet.write_string_only(3, 0, "SOUTH").unwrap();
        worksheet.write_string_only(4, 0, "South ").unwrap();
        worksheet.write_string_only(5, 0, " South").unwrap();
        worksheet.write_string_only(6, 0, " South ").unwrap();
        worksheet.write_string_only(7, 0, "Mouth").unwrap();

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

        worksheet.write_string_only(0, 0, "Header").unwrap();
        worksheet.write_number_only(1, 0, 1000).unwrap();
        worksheet.write_number_only(2, 0, 1000.0).unwrap();
        worksheet.write_string_only(3, 0, "1000").unwrap();
        worksheet.write_string_only(4, 0, " 1000 ").unwrap();
        worksheet.write_number_only(5, 0, 2000).unwrap();

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

        let result = worksheet.set_vertical_page_breaks(&[COL_MAX as u32]);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));
    }

    #[test]
    fn set_header_image() {
        let mut worksheet = Worksheet::new();

        let image = Image::new("tests/input/images/red.jpg").unwrap();
        worksheet.set_header("&R&G");

        // Test inserting an image without a matching header position.
        let result = worksheet.set_header_image(&image, XlsxImagePosition::Left);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));
    }

    #[test]
    fn rich_string() {
        let mut worksheet = Worksheet::new();

        // Test an empty array.
        let segments = [];
        let result = worksheet.write_rich_string_only(0, 0, &segments);
        assert!(matches!(result, Err(XlsxError::ParameterError(_))));

        // Test an empty string.
        let default = Format::default();
        let segments = [(&default, "")];
        let result = worksheet.write_rich_string_only(0, 0, &segments);
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

        for test_data in formulas.iter() {
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
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:16".to_string()), (1, "17:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_2() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (1..18).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:15".to_string()), (1, "16:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_3() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (2..19).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:14".to_string()), (1, "15:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_4() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (3..20).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:13".to_string()), (1, "14:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_5() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (4..21).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:12".to_string()), (1, "13:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_6() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (5..22).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:11".to_string()), (1, "12:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_7() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (6..23).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:10".to_string()), (1, "11:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_8() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (7..24).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:9".to_string()), (1, "10:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_9() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (8..25).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:8".to_string()), (1, "9:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_10() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (9..26).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:7".to_string()), (1, "8:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_11() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (10..27).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:6".to_string()), (1, "7:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_12() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (11..28).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:5".to_string()), (1, "6:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_13() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (12..29).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:4".to_string()), (1, "5:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_14() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (13..30).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:3".to_string()), (1, "4:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_15() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (14..31).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:2".to_string()), (1, "3:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_16() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (15..32).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:1".to_string()), (1, "2:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_17() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (16..33).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(1, "1:16".to_string()), (2, "17:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_18() {
        let mut worksheet = Worksheet::new();

        for (col_num, row_num) in (16..33).enumerate() {
            worksheet
                .write_number_only(row_num, (col_num + 1) as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(1, "2:17".to_string()), (2, "18:18".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
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

        assert_eq!(worksheet.check_dimensions(ROW_MAX, 0), false);
        assert_eq!(worksheet.check_dimensions(0, COL_MAX), false);

        let result = worksheet.write_string(ROW_MAX, 0, "Foo", &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.write_string_only(ROW_MAX, 0, "Foo");
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.write_number(ROW_MAX, 0, 0, &format);
        assert!(matches!(result, Err(XlsxError::RowColumnLimitError)));

        let result = worksheet.write_number_only(ROW_MAX, 0, 0);
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

        let result = worksheet.write_string_only(0, 0, long_string.unwrap());
        assert!(matches!(result, Err(XlsxError::MaxStringLengthExceeded)));
    }

    #[test]
    fn dates_and_times() {
        let mut worksheet = Worksheet::new();

        // Test date and time
        let datetimes = vec![
            (1899, 12, 31, 0, 0, 0, 0, 0.0),
            (1982, 8, 25, 0, 15, 20, 213, 30188.010650613425),
            (2065, 4, 19, 0, 16, 48, 290, 60376.011670023145),
            (2147, 12, 15, 0, 55, 25, 446, 90565.038488958337),
            (2230, 8, 10, 1, 2, 46, 891, 120753.04359827546),
            (2313, 4, 6, 1, 4, 15, 597, 150942.04462496529),
            (2395, 11, 30, 1, 9, 40, 889, 181130.04838991899),
            (2478, 7, 25, 1, 11, 32, 560, 211318.04968240741),
            (2561, 3, 21, 1, 30, 19, 169, 241507.06272186342),
            (2643, 11, 15, 1, 48, 25, 580, 271695.07529606484),
            (2726, 7, 12, 2, 3, 31, 919, 301884.08578609955),
            (2809, 3, 6, 2, 11, 11, 986, 332072.09111094906),
            (2891, 10, 31, 2, 24, 37, 95, 362261.10042934027),
            (2974, 6, 26, 2, 35, 7, 220, 392449.10772245371),
            (3057, 2, 19, 2, 45, 12, 109, 422637.1147234838),
            (3139, 10, 17, 3, 6, 39, 990, 452826.12962951389),
            (3222, 6, 11, 3, 8, 8, 251, 483014.13065105322),
            (3305, 2, 5, 3, 19, 12, 576, 513203.13834),
            (3387, 10, 1, 3, 29, 42, 574, 543391.14563164348),
            (3470, 5, 27, 3, 37, 30, 813, 573579.15105107636),
            (3553, 1, 21, 4, 14, 38, 231, 603768.17683137732),
            (3635, 9, 16, 4, 16, 28, 559, 633956.17810832174),
            (3718, 5, 13, 4, 17, 58, 222, 664145.17914608796),
            (3801, 1, 6, 4, 21, 41, 794, 694333.18173372687),
            (3883, 9, 2, 4, 56, 35, 792, 724522.20596981479),
            (3966, 4, 28, 5, 25, 14, 885, 754710.2258667245),
            (4048, 12, 21, 5, 26, 5, 724, 784898.22645513888),
            (4131, 8, 18, 5, 46, 44, 68, 815087.24078782403),
            (4214, 4, 13, 5, 48, 1, 141, 845275.24167987274),
            (4296, 12, 7, 5, 53, 52, 315, 875464.24574438657),
            (4379, 8, 3, 6, 14, 48, 580, 905652.26028449077),
            (4462, 3, 28, 6, 46, 15, 738, 935840.28212659725),
            (4544, 11, 22, 7, 31, 20, 407, 966029.31343063654),
            (4627, 7, 19, 7, 58, 33, 754, 996217.33233511576),
            (4710, 3, 15, 8, 7, 43, 130, 1026406.3386936343),
            (4792, 11, 7, 8, 29, 11, 91, 1056594.3536005903),
            (4875, 7, 4, 9, 8, 15, 328, 1086783.3807329629),
            (4958, 2, 27, 9, 30, 41, 781, 1116971.3963169097),
            (5040, 10, 23, 9, 34, 4, 462, 1147159.3986627546),
            (5123, 6, 20, 9, 37, 23, 945, 1177348.4009715857),
            (5206, 2, 12, 9, 37, 56, 655, 1207536.4013501736),
            (5288, 10, 8, 9, 45, 12, 230, 1237725.406391551),
            (5371, 6, 4, 9, 54, 14, 782, 1267913.412671088),
            (5454, 1, 28, 9, 54, 22, 108, 1298101.4127558796),
            (5536, 9, 24, 10, 1, 36, 151, 1328290.4177795255),
            (5619, 5, 20, 12, 9, 48, 602, 1358478.5068125231),
            (5702, 1, 14, 12, 34, 8, 549, 1388667.5237100578),
            (5784, 9, 8, 12, 56, 6, 495, 1418855.5389640625),
            (5867, 5, 6, 12, 58, 58, 217, 1449044.5409515856),
            (5949, 12, 30, 12, 59, 54, 263, 1479232.5416002662),
            (6032, 8, 24, 13, 34, 41, 331, 1509420.5657561459),
            (6115, 4, 21, 13, 58, 28, 601, 1539609.5822754744),
            (6197, 12, 14, 14, 2, 16, 899, 1569797.5849178126),
            (6280, 8, 10, 14, 36, 17, 444, 1599986.6085352316),
            (6363, 4, 6, 14, 37, 57, 451, 1630174.60969272),
            (6445, 11, 30, 14, 57, 42, 757, 1660363.6234115392),
            (6528, 7, 26, 15, 10, 48, 307, 1690551.6325035533),
            (6611, 3, 22, 15, 14, 39, 890, 1720739.635183912),
            (6693, 11, 15, 15, 19, 47, 988, 1750928.6387498612),
            (6776, 7, 11, 16, 4, 24, 344, 1781116.6697262037),
            (6859, 3, 7, 16, 22, 23, 952, 1811305.6822216667),
            (6941, 10, 31, 16, 29, 55, 999, 1841493.6874536921),
            (7024, 6, 26, 16, 58, 20, 259, 1871681.7071789235),
            (7107, 2, 21, 17, 4, 2, 415, 1901870.7111390624),
            (7189, 10, 16, 17, 18, 29, 630, 1932058.7211762732),
            (7272, 6, 11, 17, 47, 21, 323, 1962247.7412190163),
            (7355, 2, 5, 17, 53, 29, 866, 1992435.7454845603),
            (7437, 10, 2, 17, 53, 41, 76, 2022624.7456143056),
            (7520, 5, 28, 17, 55, 6, 44, 2052812.7465977315),
            (7603, 1, 21, 18, 14, 49, 151, 2083000.7602910995),
            (7685, 9, 16, 18, 17, 45, 738, 2113189.7623349307),
            (7768, 5, 12, 18, 29, 59, 700, 2143377.7708298611),
            (7851, 1, 7, 18, 33, 21, 233, 2173566.773162419),
            (7933, 9, 2, 19, 14, 24, 673, 2203754.8016744559),
            (8016, 4, 27, 19, 17, 12, 816, 2233942.8036205554),
            (8098, 12, 22, 19, 23, 36, 418, 2264131.8080603937),
            (8181, 8, 17, 19, 46, 25, 908, 2294319.8239109721),
            (8264, 4, 13, 20, 7, 47, 314, 2324508.8387420601),
            (8346, 12, 8, 20, 31, 37, 603, 2354696.855296331),
            (8429, 8, 3, 20, 39, 57, 770, 2384885.8610853008),
            (8512, 3, 29, 20, 50, 17, 67, 2415073.8682530904),
            (8594, 11, 22, 21, 2, 57, 827, 2445261.8770581828),
            (8677, 7, 19, 21, 23, 5, 519, 2475450.8910360998),
            (8760, 3, 14, 21, 34, 49, 572, 2505638.8991848612),
            (8842, 11, 8, 21, 39, 5, 944, 2535827.9021521294),
            (8925, 7, 4, 21, 39, 18, 426, 2566015.9022965971),
            (9008, 2, 28, 21, 46, 7, 769, 2596203.9070343636),
            (9090, 10, 24, 21, 57, 55, 662, 2626392.9152275696),
            (9173, 6, 19, 22, 19, 11, 732, 2656580.9299968979),
            (9256, 2, 13, 22, 23, 51, 376, 2686769.9332335186),
            (9338, 10, 9, 22, 27, 58, 771, 2716957.9360968866),
            (9421, 6, 5, 22, 43, 30, 392, 2747146.9468795368),
            (9504, 1, 30, 22, 48, 25, 834, 2777334.9502990046),
            (9586, 9, 24, 22, 53, 51, 727, 2807522.9540709145),
            (9669, 5, 20, 23, 12, 56, 536, 2837711.9673210187),
            (9752, 1, 14, 23, 15, 54, 109, 2867899.9693762613),
            (9834, 9, 10, 23, 17, 12, 632, 2898088.9702850925),
            (9999, 12, 31, 23, 59, 59, 0, 2958465.999988426),
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
            (2400, 1, 1, 182623.0),
            (2400, 1, 31, 182653.0),
            (2400, 2, 1, 182654.0),
            (2400, 2, 29, 182682.0),
            (2400, 3, 1, 182683.0),
            (2400, 3, 31, 182713.0),
            (2400, 4, 1, 182714.0),
            (2400, 4, 30, 182743.0),
            (2400, 5, 1, 182744.0),
            (2400, 5, 31, 182774.0),
            (2400, 6, 1, 182775.0),
            (2400, 6, 30, 182804.0),
            (2400, 7, 1, 182805.0),
            (2400, 7, 31, 182835.0),
            (2400, 8, 1, 182836.0),
            (2400, 8, 31, 182866.0),
            (2400, 9, 1, 182867.0),
            (2400, 9, 30, 182896.0),
            (2400, 10, 1, 182897.0),
            (2400, 10, 31, 182927.0),
            (2400, 11, 1, 182928.0),
            (2400, 11, 30, 182957.0),
            (2400, 12, 1, 182958.0),
            (2400, 12, 31, 182988.0),
            (4000, 1, 1, 767011.0),
            (4000, 1, 31, 767041.0),
            (4000, 2, 1, 767042.0),
            (4000, 2, 29, 767070.0),
            (4000, 3, 1, 767071.0),
            (4000, 3, 31, 767101.0),
            (4000, 4, 1, 767102.0),
            (4000, 4, 30, 767131.0),
            (4000, 5, 1, 767132.0),
            (4000, 5, 31, 767162.0),
            (4000, 6, 1, 767163.0),
            (4000, 6, 30, 767192.0),
            (4000, 7, 1, 767193.0),
            (4000, 7, 31, 767223.0),
            (4000, 8, 1, 767224.0),
            (4000, 8, 31, 767254.0),
            (4000, 9, 1, 767255.0),
            (4000, 9, 30, 767284.0),
            (4000, 10, 1, 767285.0),
            (4000, 10, 31, 767315.0),
            (4000, 11, 1, 767316.0),
            (4000, 11, 30, 767345.0),
            (4000, 12, 1, 767346.0),
            (4000, 12, 31, 767376.0),
            (4321, 1, 1, 884254.0),
            (4321, 1, 31, 884284.0),
            (4321, 2, 1, 884285.0),
            (4321, 2, 28, 884312.0),
            (4321, 3, 1, 884313.0),
            (4321, 3, 31, 884343.0),
            (4321, 4, 1, 884344.0),
            (4321, 4, 30, 884373.0),
            (4321, 5, 1, 884374.0),
            (4321, 5, 31, 884404.0),
            (4321, 6, 1, 884405.0),
            (4321, 6, 30, 884434.0),
            (4321, 7, 1, 884435.0),
            (4321, 7, 31, 884465.0),
            (4321, 8, 1, 884466.0),
            (4321, 8, 31, 884496.0),
            (4321, 9, 1, 884497.0),
            (4321, 9, 30, 884526.0),
            (4321, 10, 1, 884527.0),
            (4321, 10, 31, 884557.0),
            (4321, 11, 1, 884558.0),
            (4321, 11, 30, 884587.0),
            (4321, 12, 1, 884588.0),
            (4321, 12, 31, 884618.0),
            (9999, 1, 1, 2958101.0),
            (9999, 1, 31, 2958131.0),
            (9999, 2, 1, 2958132.0),
            (9999, 2, 28, 2958159.0),
            (9999, 3, 1, 2958160.0),
            (9999, 3, 31, 2958190.0),
            (9999, 4, 1, 2958191.0),
            (9999, 4, 30, 2958220.0),
            (9999, 5, 1, 2958221.0),
            (9999, 5, 31, 2958251.0),
            (9999, 6, 1, 2958252.0),
            (9999, 6, 30, 2958281.0),
            (9999, 7, 1, 2958282.0),
            (9999, 7, 31, 2958312.0),
            (9999, 8, 1, 2958313.0),
            (9999, 8, 31, 2958343.0),
            (9999, 9, 1, 2958344.0),
            (9999, 9, 30, 2958373.0),
            (9999, 10, 1, 2958374.0),
            (9999, 10, 31, 2958404.0),
            (9999, 11, 1, 2958405.0),
            (9999, 11, 30, 2958434.0),
            (9999, 12, 1, 2958435.0),
            (9999, 12, 31, 2958465.0),
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
        let times = vec![
            (0, 0, 0, 0, 0.0),
            (0, 15, 20, 213, 1.0650613425925924E-2),
            (0, 16, 48, 290, 1.1670023148148148E-2),
            (0, 55, 25, 446, 3.8488958333333337E-2),
            (1, 2, 46, 891, 4.3598275462962965E-2),
            (1, 4, 15, 597, 4.4624965277777782E-2),
            (1, 9, 40, 889, 4.8389918981481483E-2),
            (1, 11, 32, 560, 4.9682407407407404E-2),
            (1, 30, 19, 169, 6.2721863425925936E-2),
            (1, 48, 25, 580, 7.5296064814814809E-2),
            (2, 3, 31, 919, 8.5786099537037031E-2),
            (2, 11, 11, 986, 9.1110949074074077E-2),
            (2, 24, 37, 95, 0.10042934027777778),
            (2, 35, 7, 220, 0.1077224537037037),
            (2, 45, 12, 109, 0.11472348379629631),
            (3, 6, 39, 990, 0.12962951388888888),
            (3, 8, 8, 251, 0.13065105324074075),
            (3, 19, 12, 576, 0.13833999999999999),
            (3, 29, 42, 574, 0.14563164351851851),
            (3, 37, 30, 813, 0.1510510763888889),
            (4, 14, 38, 231, 0.1768313773148148),
            (4, 16, 28, 559, 0.17810832175925925),
            (4, 17, 58, 222, 0.17914608796296297),
            (4, 21, 41, 794, 0.18173372685185185),
            (4, 56, 35, 792, 0.2059698148148148),
            (5, 25, 14, 885, 0.22586672453703704),
            (5, 26, 5, 724, 0.22645513888888891),
            (5, 46, 44, 68, 0.24078782407407406),
            (5, 48, 1, 141, 0.2416798726851852),
            (5, 53, 52, 315, 0.24574438657407408),
            (6, 14, 48, 580, 0.26028449074074073),
            (6, 46, 15, 738, 0.28212659722222222),
            (7, 31, 20, 407, 0.31343063657407405),
            (7, 58, 33, 754, 0.33233511574074076),
            (8, 7, 43, 130, 0.33869363425925925),
            (8, 29, 11, 91, 0.35360059027777774),
            (9, 8, 15, 328, 0.380732962962963),
            (9, 30, 41, 781, 0.39631690972222228),
            (9, 34, 4, 462, 0.39866275462962958),
            (9, 37, 23, 945, 0.40097158564814817),
            (9, 37, 56, 655, 0.40135017361111114),
            (9, 45, 12, 230, 0.40639155092592594),
            (9, 54, 14, 782, 0.41267108796296298),
            (9, 54, 22, 108, 0.41275587962962962),
            (10, 1, 36, 151, 0.41777952546296299),
            (12, 9, 48, 602, 0.50681252314814818),
            (12, 34, 8, 549, 0.52371005787037039),
            (12, 56, 6, 495, 0.53896406249999995),
            (12, 58, 58, 217, 0.54095158564814816),
            (12, 59, 54, 263, 0.54160026620370372),
            (13, 34, 41, 331, 0.56575614583333333),
            (13, 58, 28, 601, 0.58227547453703699),
            (14, 2, 16, 899, 0.58491781249999997),
            (14, 36, 17, 444, 0.60853523148148148),
            (14, 37, 57, 451, 0.60969271990740748),
            (14, 57, 42, 757, 0.6234115393518519),
            (15, 10, 48, 307, 0.6325035532407407),
            (15, 14, 39, 890, 0.63518391203703706),
            (15, 19, 47, 988, 0.63874986111111109),
            (16, 4, 24, 344, 0.66972620370370362),
            (16, 22, 23, 952, 0.68222166666666662),
            (16, 29, 55, 999, 0.6874536921296297),
            (16, 58, 20, 259, 0.70717892361111112),
            (17, 4, 2, 415, 0.71113906250000003),
            (17, 18, 29, 630, 0.72117627314814825),
            (17, 47, 21, 323, 0.74121901620370367),
            (17, 53, 29, 866, 0.74548456018518516),
            (17, 53, 41, 76, 0.74561430555555563),
            (17, 55, 6, 44, 0.74659773148148145),
            (18, 14, 49, 151, 0.760291099537037),
            (18, 17, 45, 738, 0.76233493055555546),
            (18, 29, 59, 700, 0.77082986111111118),
            (18, 33, 21, 233, 0.77316241898148153),
            (19, 14, 24, 673, 0.80167445601851861),
            (19, 17, 12, 816, 0.80362055555555545),
            (19, 23, 36, 418, 0.80806039351851855),
            (19, 46, 25, 908, 0.82391097222222232),
            (20, 7, 47, 314, 0.83874206018518516),
            (20, 31, 37, 603, 0.85529633101851854),
            (20, 39, 57, 770, 0.86108530092592594),
            (20, 50, 17, 67, 0.86825309027777775),
            (21, 2, 57, 827, 0.87705818287037041),
            (21, 23, 5, 519, 0.891036099537037),
            (21, 34, 49, 572, 0.89918486111111118),
            (21, 39, 5, 944, 0.90215212962962965),
            (21, 39, 18, 426, 0.90229659722222222),
            (21, 46, 7, 769, 0.90703436342592603),
            (21, 57, 55, 662, 0.91522756944444439),
            (22, 19, 11, 732, 0.92999689814814823),
            (22, 23, 51, 376, 0.93323351851851843),
            (22, 27, 58, 771, 0.93609688657407408),
            (22, 43, 30, 392, 0.94687953703703709),
            (22, 48, 25, 834, 0.95029900462962968),
            (22, 53, 51, 727, 0.95407091435185187),
            (23, 12, 56, 536, 0.96732101851851848),
            (23, 15, 54, 109, 0.96937626157407408),
            (23, 17, 12, 632, 0.97028509259259266),
            (23, 59, 59, 999, 0.99999998842592586),
        ];

        for test_data in times {
            let (hour, min, seconds, millis, expected) = test_data;
            let datetime = NaiveTime::from_hms_milli_opt(hour, min, seconds, millis).unwrap();
            let mut diff = worksheet.time_to_excel(&datetime) - expected;
            diff = diff.abs();
            assert!(diff < 0.00000000001);
        }
    }
}
