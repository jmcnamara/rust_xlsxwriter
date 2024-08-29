// Formula unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod formula_tests {
    use crate::Formula;

    #[test]
    fn test_future_function_escapes() {
        let formula_strings = vec![
            // Test simple escapes for formulas.
            ("=foo()", "foo()"),
            ("{foo()}", "foo()"),
            ("{=foo()}", "foo()"),
            // Dynamic functions.
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
            // Newer dynamic functions (some duplicates with above).
            ("BYCOL(E1:G2)", "_xlfn.BYCOL(E1:G2)"),
            ("BYROW(E1:G2)", "_xlfn.BYROW(E1:G2)"),
            ("CHOOSECOLS(E1:G2,1)", "_xlfn.CHOOSECOLS(E1:G2,1)"),
            ("CHOOSEROWS(E1:G2,1)", "_xlfn.CHOOSEROWS(E1:G2,1)"),
            ("DROP(E1:G2,1)", "_xlfn.DROP(E1:G2,1)"),
            ("EXPAND(E1:G2,2)", "_xlfn.EXPAND(E1:G2,2)"),
            ("FILTER(E1:G2,H1:H2)", "_xlfn._xlws.FILTER(E1:G2,H1:H2)"),
            ("HSTACK(E1:G2)", "_xlfn.HSTACK(E1:G2)"),
            (
                "LAMBDA(_xlpm.number, _xlpm.number + 1)",
                "_xlfn.LAMBDA(_xlpm.number, _xlpm.number + 1)",
            ),
            (
                "MAKEARRAY(1,1,LAMBDA(_xlpm.row,_xlpm.col,TRUE)",
                "_xlfn.MAKEARRAY(1,1,_xlfn.LAMBDA(_xlpm.row,_xlpm.col,TRUE)",
            ),
            ("MAP(E1:G2,LAMBDA()", "_xlfn.MAP(E1:G2,_xlfn.LAMBDA()"),
            ("RANDARRAY(1)", "_xlfn.RANDARRAY(1)"),
            (
                "REDUCE(\"1,2,3\",E1:G2,LAMBDA()",
                "_xlfn.REDUCE(\"1,2,3\",E1:G2,_xlfn.LAMBDA()",
            ),
            (
                "SCAN(\"1,2,3\",E1:G2,LAMBDA()",
                "_xlfn.SCAN(\"1,2,3\",E1:G2,_xlfn.LAMBDA()",
            ),
            ("SEQUENCE(E1:E2)", "_xlfn.SEQUENCE(E1:E2)"),
            ("SORT(F1)", "_xlfn._xlws.SORT(F1)"),
            ("SORTBY(E1:G1,E2:G2)", "_xlfn.SORTBY(E1:G1,E2:G2)"),
            ("SWITCH(WEEKDAY(E1)", "_xlfn.SWITCH(WEEKDAY(E1)"),
            ("TAKE(E1:G2,1)", "_xlfn.TAKE(E1:G2,1)"),
            (
                "TEXTSPLIT(\"foo bar\", \" \")",
                "_xlfn.TEXTSPLIT(\"foo bar\", \" \")",
            ),
            ("TOCOL(E1:G1)", "_xlfn.TOCOL(E1:G1)"),
            ("TOROW(E1:E2)", "_xlfn.TOROW(E1:E2)"),
            ("UNIQUE(E1:G1)", "_xlfn.UNIQUE(E1:G1)"),
            ("VSTACK(E1:G2)", "_xlfn.VSTACK(E1:G2)"),
            ("WRAPCOLS(E1:F1,2)", "_xlfn.WRAPCOLS(E1:F1,2)"),
            ("WRAPROWS(E1:F1,2)", "_xlfn.WRAPROWS(E1:F1,2)"),
            (
                "XLOOKUP(M34,I35:I42,J35:K42)",
                "_xlfn.XLOOKUP(M34,I35:I42,J35:K42)",
            ),
            // Future functions.
            ("COT()", "_xlfn.COT()"),
            ("CSC()", "_xlfn.CSC()"),
            ("IFS()", "_xlfn.IFS()"),
            ("LET()", "_xlfn.LET()"),
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
            ("IMAGE()", "_xlfn.IMAGE()"),
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
            ("XMATCH()", "_xlfn.XMATCH()"),
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
            ("ISOMITTED()", "_xlfn.ISOMITTED()"),
            ("TEXTAFTER()", "_xlfn.TEXTAFTER()"),
            ("BINOM.DIST()", "_xlfn.BINOM.DIST()"),
            ("CHISQ.DIST()", "_xlfn.CHISQ.DIST()"),
            ("CHISQ.TEST()", "_xlfn.CHISQ.TEST()"),
            ("EXPON.DIST()", "_xlfn.EXPON.DIST()"),
            ("FLOOR.MATH()", "_xlfn.FLOOR.MATH()"),
            ("GAMMA.DIST()", "_xlfn.GAMMA.DIST()"),
            ("ISOWEEKNUM()", "_xlfn.ISOWEEKNUM()"),
            ("NORM.S.INV()", "_xlfn.NORM.S.INV()"),
            ("WEBSERVICE()", "_xlfn.WEBSERVICE()"),
            ("TEXTBEFORE()", "_xlfn.TEXTBEFORE()"),
            ("ERF.PRECISE()", "_xlfn.ERF.PRECISE()"),
            ("FORMULATEXT()", "_xlfn.FORMULATEXT()"),
            ("LOGNORM.INV()", "_xlfn.LOGNORM.INV()"),
            ("NORM.S.DIST()", "_xlfn.NORM.S.DIST()"),
            ("NUMBERVALUE()", "_xlfn.NUMBERVALUE()"),
            ("QUERYSTRING()", "_xlfn.QUERYSTRING()"),
            ("ARRAYTOTEXT()", "_xlfn.ARRAYTOTEXT()"),
            ("VALUETOTEXT()", "_xlfn.VALUETOTEXT()"),
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

        for &(formula_string, expected_formula) in &formula_strings {
            let formula = Formula::new(formula_string);
            assert_eq!(formula.formula_string, expected_formula);
        }
    }
}

#[test]
fn test_parse_function_escapes() {
    use crate::Formula;
    // Slightly more rigorous tests of function edge case.
    let formula_strings = vec![
        // Test simple escapes for formulas.
        ("", "", false),
        ("FOO()", "FOO()", false),
        ("=FOO()", "FOO()", false),
        ("{=FOO()}", "FOO()", false),
        ("FOO() + A1", "FOO() + A1", false),
        (r#""FOO()""#, r#""FOO()""#, false),
        (r#""FOO() 😀""#, r#""FOO() 😀""#, false),
        (r#" "😀" & FOO()"#, r#" "😀" & FOO()"#, false),
        (r#""FOO()"&"BAR()""#, r#""FOO()"&"BAR()""#, false),
        (r#"""""&FOO()"#, r#"""""&FOO()"#, false),
        ("SEQUENCE(10)", "_xlfn.SEQUENCE(10)", true),
        (
            r#"=INDEX(C:C,MATCH(MINIFS(A:A,A:A,">="&EDATE(A5,1)),A:A))"#,
            r#"INDEX(C:C,MATCH(_xlfn.MINIFS(A:A,A:A,">="&EDATE(A5,1)),A:A))"#,
            false,
        ),
        ("POISSON.DIST(A1:A3)", "_xlfn.POISSON.DIST(A1:A3)", false),
        (
            "SEQUENCE(10) + POISSON.DIST(A1:A3)",
            "_xlfn.SEQUENCE(10) + _xlfn.POISSON.DIST(A1:A3)",
            true,
        ),
        (
            "_xlfn.SEQUENCE(10) + _xlfn.POISSON.DIST(A1:A3)",
            "_xlfn.SEQUENCE(10) + _xlfn.POISSON.DIST(A1:A3)",
            true,
        ),
        ("FILTER(A1:A10,1)", "_xlfn._xlws.FILTER(A1:A10,1)", true),
    ];

    for &(input_string, expected_formula, expected_dynamic) in &formula_strings {
        let formula = Formula::new(input_string);
        assert_eq!(formula.formula_string, expected_formula);
        assert_eq!(formula.has_dynamic_function, expected_dynamic);
    }
}

#[test]
fn test_escape_table_functions() {
    use crate::Formula;
    // Test table escapes.
    let formula_strings = vec![
        ("", ""),
        ("@", "[#This Row],"),
        (r#""@""#, r#""@""#),
        (
            "SUM(Table1[@[Column1]:[Column3]])",
            "SUM(Table1[[#This Row],[Column1]:[Column3]])",
        ),
        (
            r#"=HYPERLINK(CONCAT("http://myweb.com:1677/'@md=d&path/to/sour/...'@/",[@CL],"?ac=10"),[@CL])"#,
            r#"HYPERLINK(_xlfn.CONCAT("http://myweb.com:1677/'@md=d&path/to/sour/...'@/",[[#This Row],CL],"?ac=10"),[[#This Row],CL])"#,
        ),
    ];

    for &(input_string, expected_formula) in &formula_strings {
        let formula = Formula::new(input_string).escape_table_functions();

        assert_eq!(formula.formula_string, expected_formula);
    }
}
