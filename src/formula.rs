// format - A module for representing Excel worksheet formulas.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::borrow::Cow;

use regex::Regex;

/// TODO
pub struct Formula {
    formula_string: String,
    expand_future_functions: bool,
    pub(crate) result: Box<str>,
}

impl Formula {
    /// TODO
    pub fn new(formula: impl Into<String>) -> Formula {
        Formula {
            formula_string: formula.into(),
            expand_future_functions: false,
            result: Box::from(""),
        }
    }

    /// Todo
    pub fn set_result(mut self, result: impl Into<String>) -> Formula {
        self.result = Box::from(result.into());
        self
    }

    // Check of a dynamic function/formula.
    pub(crate) fn is_dynamic_function(&self) -> bool {
        lazy_static! {
            static ref DYNAMIC_FUNCTION: Regex = Regex::new(
                r"\b(ANCHORARRAY|FILTER|LAMBDA|LET|RANDARRAY|SEQUENCE|SINGLE|SORTBY|SORT|UNIQUE|XLOOKUP|XMATCH)\("
            )
            .unwrap();
        }
        DYNAMIC_FUNCTION.is_match(&self.formula_string)
    }

    // Utility method to strip equal sign and array braces from a formula and
    // also expand out future and dynamic array formulas.
    pub(crate) fn expand_formula(&self, global_expand_future_functions: bool) -> Box<str> {
        let mut formula = self.formula_string.as_str();

        // Remove array formula braces and the leading = if they exist.
        if let Some(stripped) = formula.strip_prefix('{') {
            formula = stripped;
        }
        if let Some(stripped) = formula.strip_prefix('=') {
            formula = stripped;
        }
        if let Some(stripped) = formula.strip_suffix('}') {
            formula = stripped;
        }

        // Exit if formula is already expanded by the user.
        if formula.contains("_xlfn.") {
            return Box::from(formula);
        }

        // Expand dynamic formulas.
        let escaped_formula = Self::escape_dynamic_formulas1(formula);
        let escaped_formula = Self::escape_dynamic_formulas2(&escaped_formula);

        let formula = if self.expand_future_functions || global_expand_future_functions {
            Self::escape_future_functions(&escaped_formula)
        } else {
            escaped_formula
        };

        Box::from(formula)
    }

    // Escape/expand the dynamic formula _xlfn functions.
    fn escape_dynamic_formulas1(formula: &str) -> Cow<str> {
        lazy_static! {
            static ref XLFN: Regex = Regex::new(
                r"\b(ANCHORARRAY|LAMBDA|LET|RANDARRAY|SEQUENCE|SINGLE|SORTBY|UNIQUE|XLOOKUP|XMATCH)\("
            )
            .unwrap();
        }
        XLFN.replace_all(formula, "_xlfn.$1(")
    }

    // Escape/expand the dynamic formula _xlfn._xlws. functions.
    fn escape_dynamic_formulas2(formula: &str) -> Cow<str> {
        lazy_static! {
            static ref XLWS: Regex = Regex::new(r"\b(FILTER|SORT)\(").unwrap();
        }
        XLWS.replace_all(formula, "_xlfn._xlws.$1(")
    }

    // Escape/expand future/_xlfn functions.
    fn escape_future_functions(formula: &str) -> Cow<str> {
        lazy_static! {
            static ref FUTURE: Regex = Regex::new(
                r"\b(ACOTH|ACOT|AGGREGATE|ARABIC|BASE|BETA\.DIST|BETA\.INV|BINOM\.DIST\.RANGE|BINOM\.DIST|BINOM\.INV|BITAND|BITLSHIFT|BITOR|BITRSHIFT|BITXOR|CEILING\.MATH|CEILING\.PRECISE|CHISQ\.DIST\.RT|CHISQ\.DIST|CHISQ\.INV\.RT|CHISQ\.INV|CHISQ\.TEST|COMBINA|CONCAT|CONFIDENCE\.NORM|CONFIDENCE\.T|COTH|COT|COVARIANCE\.P|COVARIANCE\.S|CSCH|CSC|DAYS|DECIMAL|ERF\.PRECISE|ERFC\.PRECISE|EXPON\.DIST|F\.DIST\.RT|F\.DIST|F\.INV\.RT|F\.INV|F\.TEST|FILTERXML|FLOOR\.MATH|FLOOR\.PRECISE|FORECAST\.ETS\.CONFINT|FORECAST\.ETS\.SEASONALITY|FORECAST\.ETS\.STAT|FORECAST\.ETS|FORECAST\.LINEAR|FORMULATEXT|GAMMA\.DIST|GAMMA\.INV|GAMMALN\.PRECISE|GAMMA|GAUSS|HYPGEOM\.DIST|IFNA|IFS|IMCOSH|IMCOT|IMCSCH|IMCSC|IMSECH|IMSEC|IMSINH|IMTAN|ISFORMULA|ISOWEEKNUM|LOGNORM\.DIST|LOGNORM\.INV|MAXIFS|MINIFS|MODE\.MULT|MODE\.SNGL|MUNIT|NEGBINOM\.DIST|NORM\.DIST|NORM\.INV|NORM\.S\.DIST|NORM\.S\.INV|NUMBERVALUE|PDURATION|PERCENTILE\.EXC|PERCENTILE\.INC|PERCENTRANK\.EXC|PERCENTRANK\.INC|PERMUTATIONA|PHI|POISSON\.DIST|QUARTILE\.EXC|QUARTILE\.INC|QUERYSTRING|RANK\.AVG|RANK\.EQ|RRI|SECH|SEC|SHEETS|SHEET|SKEW\.P|STDEV\.P|STDEV\.S|SWITCH|T\.DIST\.2T|T\.DIST\.RT|T\.DIST|T\.INV\.2T|T\.INV|T\.TEST|TEXTJOIN|UNICHAR|UNICODE|VAR\.P|VAR\.S|WEBSERVICE|WEIBULL\.DIST|XOR|Z\.TEST)\("
            )
            .unwrap();
        }
        FUTURE.replace_all(formula, "_xlfn.$1(")
    }
}

impl From<&str> for Formula {
    fn from(value: &str) -> Formula {
        Formula::new(value)
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::Formula;

    #[test]
    fn test_dynamic_function_escapes() {
        let formula_strings = vec![
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

        for &(formula_string, expected) in &formula_strings {
            let formula = Formula::new(formula_string);
            let prepared_formula = formula.expand_formula(true);

            assert_eq!(prepared_formula.as_ref(), expected);
        }
    }
}
