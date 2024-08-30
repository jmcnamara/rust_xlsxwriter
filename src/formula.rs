// format - A module for representing Excel worksheet formulas.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::{collections::HashMap, sync::OnceLock};

/// The `Formula` struct is used to define a worksheet formula.
///
/// The `Formula` struct creates a formula type that can be used to write
/// worksheet formulas.
///
/// In general you would use the
/// [`Worksheet::write_formula()`](crate::Worksheet::write_formula) with a
/// string representation of the formula, like this:
///
/// ```
/// # // This code is available in examples/doc_working_with_formulas_intro.rs
/// #
/// # use rust_xlsxwriter::{Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
///      worksheet.write_formula(0, 0, "=10*B1 + C1")?;
/// #
/// #      worksheet.write_number(0, 1, 5)?;
/// #      worksheet.write_number(0, 2, 1)?;
/// #
/// #     workbook.save("formula.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// The formula will then be displayed as expected in Excel:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/working_with_formulas1.png">
///
/// In order to differentiate a formula from an ordinary string (for example
/// when storing it in a data structure) you can also represent the formula with
/// a [`Formula`] struct:
///
/// ```
/// # // This code is available in examples/doc_working_with_formulas_intro2.rs
/// #
/// # use rust_xlsxwriter::{Formula, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
///     worksheet.write_formula(0, 0, Formula::new("=10*B1 + C1"))?;
/// #
/// #     worksheet.write_number(0, 1, 5)?;
/// #     worksheet.write_number(0, 2, 1)?;
/// #
/// #     workbook.save("formula.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Using a `Formula` struct also allows you to write a formula using the
/// generic [`Worksheet::write()`](crate::Worksheet::write) method:
///
/// ```
/// # // This code is available in examples/doc_working_with_formulas_intro3.rs
/// #
/// # use rust_xlsxwriter::{Formula, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
///     worksheet.write(0, 0, Formula::new("=10*B1 + C1"))?;
/// #
/// #     worksheet.write_number(0, 1, 5)?;
/// #     worksheet.write_number(0, 2, 1)?;
/// #
/// #     workbook.save("formula.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// As shown in the examples above you can write a formula and expect to have
/// the result appear in Excel. However, there are a few potential issues and
/// differences that the user of `rust_xlsxwriter` should be aware of. These are
/// explained in the sections below.
///
/// # Formula Results
///
/// The `rust_xlsxwriter` library doesn't calculate the result of a formula and
/// instead stores the value "0" as the formula result. It then sets a global
/// flag in the XLSX file to say that all formulas and functions should be
/// recalculated when the file is opened.
///
/// This works fine with Excel and the majority of spreadsheet applications.
/// However, applications that don't have a facility to calculate formulas will
/// only display the "0" result. Examples of such applications are Excel
/// viewers, PDF converters, and some mobile device applications.
///
/// If required, it is also possible to specify the calculated result of the
/// formula using the [`Worksheet::set_formula_result()`] method or the
/// [`Formula::set_result()`](Formula::set_result) method:
///
/// ```
/// # // This code is available in examples/doc_worksheet_set_formula_result.rs
/// #
/// # use rust_xlsxwriter::{Formula, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
///     // Using the formula string syntax.
///     worksheet
///         .write_formula(0, 0, "1+1")?
///         .set_formula_result(0, 0, "2");
///
///     // Or using a Formula type.
///     worksheet.write_formula(1, 0, Formula::new("2+2").set_result("4"))?;
/// #
/// #     workbook.save("formulas.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_set_formula_result.png">
///
/// One common spreadsheet application where the formula recalculation doesn't
/// work is `LibreOffice` (see the following [issue report]). If you wish to
/// force recalculation in `LibreOffice` you can use the
/// [`Worksheet::set_formula_result_default()`] method to set the default result
/// to an empty string:
///
/// ```
/// # // This code is available in examples/doc_worksheet_set_formula_result_default.rs
/// #
/// # use rust_xlsxwriter::{Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
///     worksheet.set_formula_result_default("");
/// #
/// #     workbook.save("formulas.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// [`Worksheet::set_formula_result()`]: crate::Worksheet::set_formula_result
/// [`Worksheet::set_formula_result_default()`]:
///     crate::Worksheet::set_formula_result_default
/// [issue report]: https://bugs.documentfoundation.org/show_bug.cgi?id=144819
///
/// # Non US Excel functions and syntax
///
/// Excel stores formulas in the format of the US English version, regardless of
/// the language or locale of the end-user's version of Excel. Therefore all
/// formula function names written using `rust_xlsxwriter` must be in English.
/// In addition, formulas must be written with the US style separator/range
/// operator which is a comma (not semi-colon).
///
/// Some examples of how formulas should and shouldn't be written are shown
/// below:
///
/// ```
/// # // This code is available in examples/doc_working_with_formulas_syntax.rs
/// #
/// # use rust_xlsxwriter::{Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
///     // OK.
///     worksheet.write_formula(0, 0, "=SUM(1, 2, 3)")?;
///
///     // Semi-colon separator. Causes Excel error on file opening.
///     worksheet.write_formula(1, 0, "=SUM(1; 2; 3)")?;
///
///     // French function name. Causes Excel error on file opening.
///     worksheet.write_formula(2, 0, "=SOMME(1, 2, 3)")?;
/// #
/// #     workbook.save("formula.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// If you have a non-English version of Excel you can use the following
/// multi-lingual [Formula Translator](http://en.excel-translator.de/language/)
/// to help you convert the formula. It can also replace semi-colons with
/// commas.
///
///
/// # Dynamic Array support
///
/// In Office 365 Excel introduced the concept of "Dynamic Arrays" and new
/// functions that use them. The new functions are:
///
/// - `BYCOL`
/// - `BYROW`
/// - `CHOOSECOLS`
/// - `CHOOSEROWS`
/// - `DROP`
/// - `EXPAND`
/// - `FILTER`
/// - `HSTACK`
/// - `LAMBDA`
/// - `MAKEARRAY`
/// - `MAP`
/// - `RANDARRAY`
/// - `REDUCE`
/// - `SCAN`
/// - `SEQUENCE`
/// - `SORT`
/// - `SORTBY`
/// - `SWITCH`
/// - `TAKE`
/// - `TEXTSPLIT`
/// - `TOCOL`
/// - `TOROW`
/// - `UNIQUE`
/// - `VSTACK`
/// - `WRAPCOLS`
/// - `WRAPROWS`
/// - `XLOOKUP`
///
/// The following special case functions were also added with Dynamic Arrays:
///
/// - `SINGLE`: Explained below in [The Implicit Intersection Operator
///   `@`](#the-implicit-intersection-operator-)
/// - `ANCHORARRAY`:  Explained below in [The Spilled Range Operator
///   `#`](#the-spilled-range-operator-)
///
///
/// ## Dynamic Arrays - An introduction
///
/// Dynamic arrays in Excel are ranges of return values that can change size
/// based on the results. For example, a function such as `FILTER()` returns an
/// array of values that can vary in size depending on the the filter results
///
/// ```text
/// worksheet.write_dynamic_formula(1, 5, "=FILTER(A1:D17,C1:C17=K2)")?;
/// ```
///
/// This formula gives the results shown in the image below. The dynamic range
/// here is "F2:I5" but it can vary based on the filter criteria.
///
/// <img src="https://rustxlsxwriter.github.io/images/dynamic_arrays02.png">
///
/// It is also possible to get dynamic array behavior with older Excel
/// functions. For example, the Excel function `"=LEN(A1)"` applies to a single
/// cell and returns a single value but it can also apply to a range of cells
/// and return a range of values using an array formula like `"{=LEN(A1:A3)}"`.
/// This type of "static" array behavior is referred to as a CSE
/// (Ctrl+Shift+Enter) formula and has existed in Excel since early versions. In
/// Office 365 Excel updated and extended this behavior to create the concept of
/// dynamic arrays. In Excel 365 you can now write the previous LEN function as
/// `"=LEN(A1:A3)"` and get a dynamic range of return values:
///
/// <img src="https://rustxlsxwriter.github.io/images/intersection03.png">
///
/// The difference between the two types of array functions is explained in the
/// Microsoft documentation on [Dynamic array formulas vs. legacy CSE array
/// formulas].
///
/// In `rust_xlsxwriter` you can use the [`Worksheet::write_array_formula()`]
/// function to get a static/CSE range and
/// [`Worksheet::write_dynamic_array_formula()`] or
/// [`Worksheet::write_dynamic_formula()`] to get a dynamic range.
///
/// [`Worksheet::write_array_formula()`]: crate::Worksheet::write_array_formula
/// [`Worksheet::write_dynamic_formula()`]:
///     crate::Worksheet::write_dynamic_formula
/// [`Worksheet::write_dynamic_array_formula()`]:
///     crate::Worksheet::write_dynamic_array_formula
///
/// [Dynamic array formulas in Excel]:
///     https://exceljet.net/dynamic-array-formulas-in-excel
/// [Dynamic array formulas vs. legacy CSE array formulas]:
///     https://support.microsoft.com/en-us/office/dynamic-array-formulas-vs-legacy-cse-array-formulas-ca421f1b-fbb2-4c99-9924-df571bd4f1b4
///
/// The `worksheet.write_dynamic_array_formula()` function takes a `(first_row,
/// first_col, last_row, last_col)` cell range to define the area that the
/// formula applies to. However, since the range is dynamic this generally won't
/// be known in advance in which case you can specify the range with the same
/// start and end cell. The following range is "F2:F2":
///
/// ```text
///     worksheet.write_dynamic_array_formula(1, 5, 1, 5, "=FILTER(A1:D17,C1:C17=K2)")?;
/// ```
/// As a syntactic shortcut you can use the `worksheet.write_dynamic_formula()`
/// function which only requires the start cell:
///
/// ```text
///    worksheet.write_dynamic_formula(1, 5, "=FILTER(A1:D17,C1:C17=K2)")?;
/// ```
///
/// For a wider and more general introduction to dynamic arrays see the
/// following: [Dynamic array formulas in Excel].
///
///
/// ## The Implicit Intersection Operator "@"
///
/// The Implicit Intersection Operator, "@", is used by Excel 365 to indicate a
/// position in a formula that is implicitly returning a single value when a
/// range or an array could be returned.
///
/// We can see how this operator works in practice by considering the formula we
/// used in the last section: `=LEN(A1:A3)`. In Excel versions without support
/// for dynamic arrays, i.e. prior to Excel 365, this formula would operate on a
/// single value from the input range and return a single value, like the
/// following in Excel 2011:
///
/// <img src="https://rustxlsxwriter.github.io/images/intersection01.png">
///
/// There is an implicit conversion here of the range of input values, "A1:A3",
/// to a single value "A1". Since this was the default behavior of older
/// versions of Excel this conversion isn't highlighted in any way. But if you
/// open the same file in Excel 365 it will appear as follows:
///
/// <img src="https://rustxlsxwriter.github.io/images/intersection02.png">
///
/// The result of the formula is the same (this is important to note) and it
/// still operates on, and returns, a single value. However the formula now
/// contains a "@" operator to show that it is implicitly using a single value
/// from the given range.
///
/// In Excel 365, and with [`Worksheet::write_dynamic_formula()`] in
/// `rust_xlsxwriter`, it would operate on the entire range and return an array
/// of values:
///
/// <img src="https://rustxlsxwriter.github.io/images/intersection03.png">
///
/// If you are encountering the Implicit Intersection Operator "@" for the first
/// time then it is probably from a point of view of "why is Excel or
/// `rust_xlsxwriter` putting @s in my formulas". In practical terms if you
/// encounter this operator, and you don't intend it to be there, then you
/// should probably write the formula as a CSE or dynamic array function using
/// [`Worksheet::write_array_formula()`] or
/// [`Worksheet::write_dynamic_array_formula()`]
///
///
/// A full explanation of this operator is given in the Microsoft documentation
/// on the [Implicit intersection operator: @].
///
/// [Implicit intersection operator: @]:
///     https://support.microsoft.com/en-us/office/implicit-intersection-operator-ce3be07b-0101-4450-a24e-c1c999be2b34?ui=en-us&rs=en-us&ad=us>
///
/// One important thing to note is that the "@" operator isn't stored with the
/// formula. It is just displayed by Excel 365 when reading "legacy" formulas.
/// However, it is possible to write it to a formula, if necessary, using
/// `SINGLE()`. The rare cases where this may be necessary are shown in the
/// linked document in the previous paragraph.
///
///
/// ## The Spilled Range Operator "#"
///
/// In the sections above we saw that dynamic array formulas can return variable
/// sized ranges of results. The Excel documentation refers to this as a
/// "Spilled" range/array from the idea that the results spill into the required
/// number of cells. This is explained in the Microsoft documentation on
/// [Dynamic array formulas and spilled array behavior].
///
/// [Dynamic array formulas and spilled array behavior]:
///     https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531
///
///
/// Since a spilled range is variable in size a new operator is required to
/// refer to the range. This operator is the [Spilled range operator] and it is
/// represented by "#". For example, the range `F2#` in the image below is used
/// to refer to a dynamic array returned by `UNIQUE()` in the cell `F2`:
///
/// [Spilled range operator]:
///     https://support.microsoft.com/en-us/office/spilled-range-operator-3dd5899f-bca2-4b9d-a172-3eae9ac22efd
///
/// <img src="https://rustxlsxwriter.github.io/images/spill01.png">
///
/// Unfortunately, Excel doesn't store the operator in the formula like this and
/// in `rust_xlsxwriter` you need to use the explicit function `ANCHORARRAY()`
/// to refer to a spilled range. The example in the image above was generated
/// using the following formula:
///
/// ```text
///     worksheet.write_dynamic_formula(1, 9, "=COUNTA(ANCHORARRAY(F2))")?;
/// ```
///
/// ## The Excel 365 `LAMBDA()` function
///
/// Recent versions of Excel 365 have introduced a powerful new function/feature
/// called `LAMBDA()`. This is similar to closure expressions in Rust or [lambda
/// expressions] in C++ (and other languages).
///
/// [lambda expressions]:
///     https://docs.microsoft.com/en-us/cpp/cpp/lambda-expressions-in-cpp?view=msvc-160
///
///
/// Consider the following Excel example which converts the variable `temp` from
/// Fahrenheit to Celsius:
///
/// ```text
///     LAMBDA(temp, (5/9) * (temp-32))
/// ```
///
/// This could be called in Excel with an argument:
///
/// ```text
///     =LAMBDA(temp, (5/9) * (temp-32))(212)
/// ```
///
/// Or assigned to a defined name and called as a user defined function:
///
/// ```text
///     =ToCelsius(212)
/// ```
///
/// A `rust_xlsxwriter` example that replicates the described Excel
/// functionality is shown below:
///
///
/// ```
/// # // This code is available in examples/app_lambda.rs
/// #
/// use rust_xlsxwriter::{Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Write a Lambda function to convert Fahrenheit to Celsius to a cell as a
///     // defined name and use that to calculate a value.
///     //
///     // Note that the formula name is prefixed with "_xlfn." (this is normally
///     // converted automatically by write_formula*() but isn't for defined names)
///     // and note that the lambda function parameters are prefixed with "_xlpm.".
///     // These prefixes won't show up in Excel.
///     workbook.define_name(
///         "ToCelsius",
///         "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))",
///     )?;
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Write the same Lambda function as a cell formula.
///     //
///     // Note that the lambda function parameters must be prefixed with "_xlpm.".
///     // These prefixes won't show up in Excel.
///     worksheet.write_formula(0, 0, "=LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)")?;
///
///     // The user defined name needs to be written explicitly as a dynamic array
///     // formula.
///     worksheet.write_dynamic_formula(1, 0, "=ToCelsius(212)")?;
///
///     // Save the file to disk.
///     workbook.save("lambda.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Note, that the formula name must have a `_xlfn.` prefix and the parameters
/// in the `LAMBDA()` function must have a `_xlpm.`  prefix for compatibility
/// with how the formulas are stored in Excel. These prefixes won't show up in
/// the formula, as shown in the image below.
///
/// <img src="https://rustxlsxwriter.github.io/images/app_lambda.png">
///
/// The `LET()` function is often used in conjunction with `LAMBDA()` to assign
/// names to calculation results.
///
///
/// # Formulas added in Excel 2010 and later
///
/// Excel 2010 and later versions added functions which weren't defined in the
/// original file specification. These functions are referred to by Microsoft as
/// "Future Functions". Examples of these functions are `ACOT`, `CHISQ.DIST.RT`
/// , `CONFIDENCE.NORM`, `STDEV.P`, `STDEV.S` and `WORKDAY.INTL`.
///
/// Although these formulas are displayed as normal in Excel they are stored
/// with a prefix. For example `STDEV.S(B1:B5)` is stored in the Excel file as
/// `xlfn.STDEV.S(B1:B5)`. The `rust_xlsxwriter` crate makes these changes
/// automatically so in general you don't have to worry about this unless you
/// are dealing with features such as Lambda functions, see above. However, if
/// required you can manually prefix any required function with the `_xlfn.`
/// prefix.
///
/// For completeness the following is a list of future functions taken from [MS
/// XLSX extensions documentation on future functions].
///
/// [MS XLSX extensions documentation on future functions]:
///     http://msdn.microsoft.com/en-us/library/dd907480%28v=office.12%29.aspx
///
/// Note, the Python in Excel functions aren't simple functions and aren't
/// supported.
///
/// | Future Functions                 |
/// | -------------------------------- |
/// | `_xlfn.ACOTH`                    |
/// | `_xlfn.ACOT`                     |
/// | `_xlfn.AGGREGATE`                |
/// | `_xlfn.ARABIC`                   |
/// | `_xlfn.ARRAYTOTEXT`              |
/// | `_xlfn.BASE`                     |
/// | `_xlfn.BETA.DIST`                |
/// | `_xlfn.BETA.INV`                 |
/// | `_xlfn.BINOM.DIST.RANGE`         |
/// | `_xlfn.BINOM.DIST`               |
/// | `_xlfn.BINOM.INV`                |
/// | `_xlfn.BITAND`                   |
/// | `_xlfn.BITLSHIFT`                |
/// | `_xlfn.BITOR`                    |
/// | `_xlfn.BITRSHIFT`                |
/// | `_xlfn.BITXOR`                   |
/// | `_xlfn.CEILING.MATH`             |
/// | `_xlfn.CEILING.PRECISE`          |
/// | `_xlfn.CHISQ.DIST.RT`            |
/// | `_xlfn.CHISQ.DIST`               |
/// | `_xlfn.CHISQ.INV.RT`             |
/// | `_xlfn.CHISQ.INV`                |
/// | `_xlfn.CHISQ.TEST`               |
/// | `_xlfn.COMBINA`                  |
/// | `_xlfn.CONCAT`                   |
/// | `_xlfn.CONFIDENCE.NORM`          |
/// | `_xlfn.CONFIDENCE.T`             |
/// | `_xlfn.COTH`                     |
/// | `_xlfn.COT`                      |
/// | `_xlfn.COVARIANCE.P`             |
/// | `_xlfn.COVARIANCE.S`             |
/// | `_xlfn.CSCH`                     |
/// | `_xlfn.CSC`                      |
/// | `_xlfn.DAYS`                     |
/// | `_xlfn.DECIMAL`                  |
/// | `ECMA.CEILING`                   |
/// | `_xlfn.ERF.PRECISE`              |
/// | `_xlfn.ERFC.PRECISE`             |
/// | `_xlfn.EXPON.DIST`               |
/// | `_xlfn.F.DIST.RT`                |
/// | `_xlfn.F.DIST`                   |
/// | `_xlfn.F.INV.RT`                 |
/// | `_xlfn.F.INV`                    |
/// | `_xlfn.F.TEST`                   |
/// | `_xlfn.FIELDVALUE`               |
/// | `_xlfn.FILTERXML`                |
/// | `_xlfn.FLOOR.MATH`               |
/// | `_xlfn.FLOOR.PRECISE`            |
/// | `_xlfn.FORECAST.ETS.CONFINT`     |
/// | `_xlfn.FORECAST.ETS.SEASONALITY` |
/// | `_xlfn.FORECAST.ETS.STAT`        |
/// | `_xlfn.FORECAST.ETS`             |
/// | `_xlfn.FORECAST.LINEAR`          |
/// | `_xlfn.FORMULATEXT`              |
/// | `_xlfn.GAMMA.DIST`               |
/// | `_xlfn.GAMMA.INV`                |
/// | `_xlfn.GAMMALN.PRECISE`          |
/// | `_xlfn.GAMMA`                    |
/// | `_xlfn.GAUSS`                    |
/// | `_xlfn.HYPGEOM.DIST`             |
/// | `_xlfn.IFNA`                     |
/// | `_xlfn.IFS`                      |
/// | `_xlfn.IMAGE`                    |
/// | `_xlfn.IMCOSH`                   |
/// | `_xlfn.IMCOT`                    |
/// | `_xlfn.IMCSCH`                   |
/// | `_xlfn.IMCSC`                    |
/// | `_xlfn.IMSECH`                   |
/// | `_xlfn.IMSEC`                    |
/// | `_xlfn.IMSINH`                   |
/// | `_xlfn.IMTAN`                    |
/// | `_xlfn.ISFORMULA`                |
/// | `ISO.CEILING`                    |
/// | `_xlfn.ISOMITTED`                |
/// | `_xlfn.ISOWEEKNUM`               |
/// | `_xlfn.LET`                      |
/// | `_xlfn.LOGNORM.DIST`             |
/// | `_xlfn.LOGNORM.INV`              |
/// | `_xlfn.MAXIFS`                   |
/// | `_xlfn.MINIFS`                   |
/// | `_xlfn.MODE.MULT`                |
/// | `_xlfn.MODE.SNGL`                |
/// | `_xlfn.MUNIT`                    |
/// | `_xlfn.NEGBINOM.DIST`            |
/// | `NETWORKDAYS.INTL`               |
/// | `_xlfn.NORM.DIST`                |
/// | `_xlfn.NORM.INV`                 |
/// | `_xlfn.NORM.S.DIST`              |
/// | `_xlfn.NORM.S.INV`               |
/// | `_xlfn.NUMBERVALUE`              |
/// | `_xlfn.PDURATION`                |
/// | `_xlfn.PERCENTILE.EXC`           |
/// | `_xlfn.PERCENTILE.INC`           |
/// | `_xlfn.PERCENTRANK.EXC`          |
/// | `_xlfn.PERCENTRANK.INC`          |
/// | `_xlfn.PERMUTATIONA`             |
/// | `_xlfn.PHI`                      |
/// | `_xlfn.POISSON.DIST`             |
/// | `_xlfn.PQSOURCE`                 |
/// | `_xlfn.PYTHON_STR`               |
/// | `_xlfn.PYTHON_TYPE`              |
/// | `_xlfn.PYTHON_TYPENAME`          |
/// | `_xlfn.QUARTILE.EXC`             |
/// | `_xlfn.QUARTILE.INC`             |
/// | `_xlfn.QUERYSTRING`              |
/// | `_xlfn.RANK.AVG`                 |
/// | `_xlfn.RANK.EQ`                  |
/// | `_xlfn.RRI`                      |
/// | `_xlfn.SECH`                     |
/// | `_xlfn.SEC`                      |
/// | `_xlfn.SHEETS`                   |
/// | `_xlfn.SHEET`                    |
/// | `_xlfn.SKEW.P`                   |
/// | `_xlfn.STDEV.P`                  |
/// | `_xlfn.STDEV.S`                  |
/// | `_xlfn.T.DIST.2T`                |
/// | `_xlfn.T.DIST.RT`                |
/// | `_xlfn.T.DIST`                   |
/// | `_xlfn.T.INV.2T`                 |
/// | `_xlfn.T.INV`                    |
/// | `_xlfn.T.TEST`                   |
/// | `_xlfn.TEXTAFTER`                |
/// | `_xlfn.TEXTBEFORE`               |
/// | `_xlfn.TEXTJOIN`                 |
/// | `_xlfn.UNICHAR`                  |
/// | `_xlfn.UNICODE`                  |
/// | `_xlfn.VALUETOTEXT`              |
/// | `_xlfn.VAR.P`                    |
/// | `_xlfn.VAR.S`                    |
/// | `_xlfn.WEBSERVICE`               |
/// | `_xlfn.WEIBULL.DIST`             |
/// | `WORKDAY.INTL`                   |
/// | `_xlfn.XMATCH`                   |
/// | `_xlfn.XOR`                      |
/// | `_xlfn.Z.TEST`                   |
///
/// The dynamic array functions shown in the [Dynamic Array
/// support](#dynamic-array-support) section are also future functions:
///
/// | Dynamic Array Functions          |
/// | -------------------------------- |
/// | `_xlfn.ANCHORARRAY`              |
/// | `_xlfn.BYCOL`                    |
/// | `_xlfn.BYROW`                    |
/// | `_xlfn.CHOOSECOLS`               |
/// | `_xlfn.CHOOSEROWS`               |
/// | `_xlfn.DROP`                     |
/// | `_xlfn.EXPAND`                   |
/// | `_xlfn._xlws.FILTER`             |
/// | `_xlfn.HSTACK`                   |
/// | `_xlfn.LAMBDA`                   |
/// | `_xlfn.MAKEARRAY`                |
/// | `_xlfn.MAP`                      |
/// | `_xlfn._xlws.PY`                 |
/// | `_xlfn.RANDARRAY`                |
/// | `_xlfn.REDUCE`                   |
/// | `_xlfn.SCAN`                     |
/// | `_xlfn.SINGLE`                   |
/// | `_xlfn.SEQUENCE`                 |
/// | `_xlfn._xlws.SORT`               |
/// | `_xlfn.SORTBY`                   |
/// | `_xlfn.SWITCH`                   |
/// | `_xlfn.TAKE`                     |
/// | `_xlfn.TEXTSPLIT`                |
/// | `_xlfn.TOCOL`                    |
/// | `_xlfn.TOROW`                    |
/// | `_xlfn.UNIQUE`                   |
/// | `_xlfn.VSTACK`                   |
/// | `_xlfn.WRAPCOLS`                 |
/// | `_xlfn.WRAPROWS`                 |
/// | `_xlfn.XLOOKUP`                  |
///
/// # Dealing with formula errors
///
/// If there is an error in the syntax of a formula it is usually displayed in
/// Excel as `#NAME?`. Alternatively you may get a warning from Excel when the
/// file is loaded. If you encounter an error like this you can debug it using
/// the following steps:
///
/// 1. Ensure the formula is valid in Excel by copying and pasting it into a
///    cell. Note, this should be done in Excel and **not** other applications
///    such as `OpenOffice` or `LibreOffice` since they may have slightly
///    different syntax.
///
/// 2. Ensure the formula is using comma separators instead of semi-colons, see
///    [Non US Excel functions and syntax](#non-us-excel-functions-and-syntax).
///
/// 3. Ensure the formula is in English, see [Non US Excel functions and
///    syntax](#non-us-excel-functions-and-syntax).
///
/// 4. Ensure that the formula doesn't contain an Excel 2010+ future function,
///    see [Formulas added in Excel 2010 and
///    later](#formulas-added-in-excel-2010-and-later). If it does then ensure
///    that the correct prefix is used.
///
/// 5. If the function loads in Excel but appears with one or more `@` symbols
///    added then it is probably an array function and should be written using
///    [`Worksheet::write_array_formula()`] or
///    [`Worksheet::write_dynamic_array_formula()`] (see also [Dynamic Array
///    support](#dynamic-array-support)).
///
/// Finally if you have completed all the previous steps and still get a
/// `#NAME?` error you can examine a valid Excel file to see what the correct
/// syntax should be. To do this you should create a valid formula in Excel and
/// save the file. You can then examine the XML in the unzipped file.
///
/// The following shows how to do that using Linux `unzip` and libxml's
/// [xmllint](http://xmlsoft.org/xmllint.html) to format the XML for clarity:
///
/// ```bash
///     $ unzip myfile.xlsx -d myfile
///     $ xmllint --format myfile/xl/worksheets/sheet1.xml | grep '</f>'
///
///             <f>SUM(1, 2, 3)</f>
/// ```
///
#[derive(Clone, PartialEq)]
pub struct Formula {
    pub(crate) formula_string: String,
    pub(crate) has_dynamic_function: bool,
    pub(crate) result: Box<str>,
    expand_future_functions: bool,
    expand_table_functions: bool,
}

impl Formula {
    /// Create a new `Formula` struct instance.
    ///
    /// # Parameters
    ///
    /// `formula` - A string like type representing an Excel formula.
    ///
    pub fn new(formula: impl AsRef<str>) -> Formula {
        // Remove array formula braces and the leading = if they exist.
        let mut formula = formula.as_ref();
        if let Some(stripped) = formula.strip_prefix('{') {
            formula = stripped;
        }
        if let Some(stripped) = formula.strip_prefix('=') {
            formula = stripped;
        }
        if let Some(stripped) = formula.strip_suffix('}') {
            formula = stripped;
        }

        // We need to escape future functions in a formula string. If the user
        // has already done this we simply copy the string. In both cases we
        // need to determine if it contains dynamic functions.
        let (formula_string, has_dynamic_function) = if formula.contains("_xlfn.") {
            // Already escaped.
            Self::copy_escaped_formula(formula)
        } else {
            // Needs escaping.
            Self::escape_formula(formula)
        };

        Formula {
            formula_string,
            has_dynamic_function,
            result: Box::from(""),
            expand_future_functions: false,
            expand_table_functions: false,
        }
    }

    /// Specify the result of a formula.
    ///
    /// As explained above in the section on [Formula
    /// Results](#formula-results) it is occasionally necessary to specify the
    /// result of a formula. This can be done using the `set_result()` method.
    ///
    /// # Parameters
    ///
    /// `result` - The formula result, as a string or string like type.
    ///
    /// # Examples
    ///
    /// The following example demonstrates manually setting the result of a
    /// formula. Note, this is only required for non-Excel applications that
    /// don't calculate formula results.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_formula_result.rs
    /// #
    /// # use rust_xlsxwriter::{Formula, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Using string syntax.
    ///     worksheet
    ///         .write_formula(0, 0, "1+1")?
    ///         .set_formula_result(0, 0, "2");
    ///
    ///     // Or using a Formula type.
    ///     worksheet.write_formula(1, 0, Formula::new("2+2").set_result("4"))?;
    /// #
    /// #     workbook.save("formulas.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_formula_result.png">
    ///
    pub fn set_result(mut self, result: impl Into<String>) -> Formula {
        self.result = Box::from(result.into());
        self
    }

    // Prefix any "future" functions in a formula with "_xlfn.". We parse the
    // string to avoid replacements in string literal within the formula.
    fn escape_formula(formula: &str) -> (String, bool) {
        let mut start_position = 0;
        let mut in_function = false;
        let mut in_string_literal = false;
        let mut has_dynamic_function = false;
        let mut escaped_formula = String::with_capacity(formula.len());

        for (current_position, char) in formula.char_indices() {
            // Match the start/end of string literals. We track these to avoid
            // escaping function names in strings. In Excel a double quote in a
            // string literal is doubled, so this will also match escapes.
            if char == '"' {
                in_string_literal = !in_string_literal;
            }

            // Copy the string literal.
            if in_string_literal {
                escaped_formula.push(char);
                continue;
            }

            // Function names are comprised of "A-Z", "0-9" and ".".
            let is_function_char =
                char.is_ascii_uppercase() || char.is_ascii_digit() || char == '.';

            // Simple state machine where we are either accumulating possible
            // function names in a buffer for evaluation, or copying non-function
            // name characters.
            if in_function {
                if !is_function_char {
                    let token = &formula[start_position..current_position];

                    // If the first non function char is an opening bracket then we
                    // have found a function name.
                    if char == '(' {
                        // Check if function is an Excel "future" function.
                        if let Some(function_type) = Self::future_functions(token) {
                            // Add the future function prefix.
                            escaped_formula.push_str("_xlfn.");

                            // Some functions have an additional prefix.
                            if *function_type == 2 {
                                escaped_formula.push_str("_xlws.");
                            }

                            // Check if the function is "dynamic".
                            has_dynamic_function |= *function_type > 0;
                        }
                    }

                    // Copy the token, whether it is a function name or not.
                    escaped_formula.push_str(token);
                    escaped_formula.push(char);
                    in_function = false;
                }
            } else if is_function_char {
                // Match the start of a possible function name.
                start_position = current_position;
                in_function = true;
            } else {
                escaped_formula.push(char);
            }
        }

        // Clean up any trailing buffer that wasn't a function.
        if in_function {
            escaped_formula.push_str(&formula[start_position..]);
        }

        (escaped_formula, has_dynamic_function)
    }

    // This is a version of the previous escape_formula() function that only
    // checks to see if a user escaped string contains a dynamic function and
    // returns a clone of the string.
    fn copy_escaped_formula(formula: &str) -> (String, bool) {
        let mut start_position = 0;
        let mut in_function = false;
        let mut in_string_literal = false;
        let mut has_dynamic_function = false;

        for (current_position, char) in formula.char_indices() {
            // Match the start/end of string literals. We track these to avoid
            // matching function names in strings. In Excel a double quote in a
            // string literal is doubled, so this will also match escapes.
            if char == '"' {
                in_string_literal = !in_string_literal;
            }

            // Ignore the string literal.
            if in_string_literal {
                continue;
            }

            // Function names are comprised of "A-Z", "0-9" and ".".
            let is_function_char =
                char.is_ascii_uppercase() || char.is_ascii_digit() || char == '.';
            let is_function_start_char = char.is_ascii_uppercase() || char.is_ascii_digit();

            // Simple state machine where we accumulate possible function names
            // in a buffer for evaluation.
            if in_function {
                if !is_function_char {
                    let token = &formula[start_position..current_position];

                    // If the first non function char is an opening bracket then we
                    // have found a function name.
                    if char == '(' {
                        // Check if function is an Excel "future" function.
                        if let Some(function_type) = Self::future_functions(token) {
                            has_dynamic_function |= *function_type > 0;
                        }
                    }

                    in_function = false;
                }
            } else if is_function_start_char {
                // Match the start of a possible function name.
                start_position = current_position;
                in_function = true;
            }
        }

        (formula.to_string(), has_dynamic_function)
    }

    // Escape/expand table functions. This mainly involves converting Excel 2010
    // "@" table ref to 2007 "[#This Row],". We parse the string to avoid
    // replacements in string literals within the formula.
    pub(crate) fn escape_table_functions(mut self) -> Formula {
        if !self.formula_string.contains('@') {
            // No escaping required.
            return self;
        }

        let mut in_string_literal = false;
        let mut escaped_formula = String::with_capacity(self.formula_string.len());

        for char in self.formula_string.chars() {
            // Match the start/end of string literals to avoid escaping
            // references in strings.
            if char == '"' {
                in_string_literal = !in_string_literal;
            }

            // Copy the string literal.
            if in_string_literal {
                escaped_formula.push(char);
                continue;
            }

            // Replace table reference.
            if char == '@' {
                escaped_formula.push_str("[#This Row],");
            } else {
                escaped_formula.push(char);
            }
        }

        self.formula_string = escaped_formula;
        self
    }

    // This is a lookup table to match Excel "future" functions that require a
    // prefix. The types are:
    //     0 = Standard future functions.
    //     1 = Future functions that are also dynamic functions.
    //     2 = Dynamic function that require an additional prefix.
    #[allow(clippy::too_many_lines)]
    fn future_functions(function: &str) -> Option<&u8> {
        static FUTURE_FUNCTIONS: OnceLock<HashMap<&str, u8>> = OnceLock::new();
        FUTURE_FUNCTIONS
            .get_or_init(|| {
                HashMap::from([
                    // Future functions.
                    ("ACOTH", 0),
                    ("ACOT", 0),
                    ("AGGREGATE", 0),
                    ("ARABIC", 0),
                    ("ARRAYTOTEXT", 0),
                    ("BASE", 0),
                    ("BETA.DIST", 0),
                    ("BETA.INV", 0),
                    ("BINOM.DIST.RANGE", 0),
                    ("BINOM.DIST", 0),
                    ("BINOM.INV", 0),
                    ("BITAND", 0),
                    ("BITLSHIFT", 0),
                    ("BITOR", 0),
                    ("BITRSHIFT", 0),
                    ("BITXOR", 0),
                    ("CEILING.MATH", 0),
                    ("CEILING.PRECISE", 0),
                    ("CHISQ.DIST.RT", 0),
                    ("CHISQ.DIST", 0),
                    ("CHISQ.INV.RT", 0),
                    ("CHISQ.INV", 0),
                    ("CHISQ.TEST", 0),
                    ("COMBINA", 0),
                    ("CONCAT", 0),
                    ("CONFIDENCE.NORM", 0),
                    ("CONFIDENCE.T", 0),
                    ("COTH", 0),
                    ("COT", 0),
                    ("COVARIANCE.P", 0),
                    ("COVARIANCE.S", 0),
                    ("CSCH", 0),
                    ("CSC", 0),
                    ("DAYS", 0),
                    ("DECIMAL", 0),
                    ("ERF.PRECISE", 0),
                    ("ERFC.PRECISE", 0),
                    ("EXPON.DIST", 0),
                    ("F.DIST.RT", 0),
                    ("F.DIST", 0),
                    ("F.INV.RT", 0),
                    ("F.INV", 0),
                    ("F.TEST", 0),
                    ("FIELDVALUE", 0),
                    ("FILTERXML", 0),
                    ("FLOOR.MATH", 0),
                    ("FLOOR.PRECISE", 0),
                    ("FORECAST.ETS.CONFINT", 0),
                    ("FORECAST.ETS.SEASONALITY", 0),
                    ("FORECAST.ETS.STAT", 0),
                    ("FORECAST.ETS", 0),
                    ("FORECAST.LINEAR", 0),
                    ("FORMULATEXT", 0),
                    ("GAMMA.DIST", 0),
                    ("GAMMA.INV", 0),
                    ("GAMMALN.PRECISE", 0),
                    ("GAMMA", 0),
                    ("GAUSS", 0),
                    ("HYPGEOM.DIST", 0),
                    ("IFNA", 0),
                    ("IFS", 0),
                    ("IMAGE", 0),
                    ("IMCOSH", 0),
                    ("IMCOT", 0),
                    ("IMCSCH", 0),
                    ("IMCSC", 0),
                    ("IMSECH", 0),
                    ("IMSEC", 0),
                    ("IMSINH", 0),
                    ("IMTAN", 0),
                    ("ISFORMULA", 0),
                    ("ISOMITTED", 0),
                    ("ISOWEEKNUM", 0),
                    ("LET", 0),
                    ("LOGNORM.DIST", 0),
                    ("LOGNORM.INV", 0),
                    ("MAXIFS", 0),
                    ("MINIFS", 0),
                    ("MODE.MULT", 0),
                    ("MODE.SNGL", 0),
                    ("MUNIT", 0),
                    ("NEGBINOM.DIST", 0),
                    ("NORM.DIST", 0),
                    ("NORM.INV", 0),
                    ("NORM.S.DIST", 0),
                    ("NORM.S.INV", 0),
                    ("NUMBERVALUE", 0),
                    ("PDURATION", 0),
                    ("PERCENTILE.EXC", 0),
                    ("PERCENTILE.INC", 0),
                    ("PERCENTRANK.EXC", 0),
                    ("PERCENTRANK.INC", 0),
                    ("PERMUTATIONA", 0),
                    ("PHI", 0),
                    ("POISSON.DIST", 0),
                    ("PQSOURCE", 0),
                    ("PYTHON_STR", 0),
                    ("PYTHON_TYPE", 0),
                    ("PYTHON_TYPENAME", 0),
                    ("QUARTILE.EXC", 0),
                    ("QUARTILE.INC", 0),
                    ("QUERYSTRING", 0),
                    ("RANK.AVG", 0),
                    ("RANK.EQ", 0),
                    ("RRI", 0),
                    ("SECH", 0),
                    ("SEC", 0),
                    ("SHEETS", 0),
                    ("SHEET", 0),
                    ("SKEW.P", 0),
                    ("STDEV.P", 0),
                    ("STDEV.S", 0),
                    ("T.DIST.2T", 0),
                    ("T.DIST.RT", 0),
                    ("T.DIST", 0),
                    ("T.INV.2T", 0),
                    ("T.INV", 0),
                    ("T.TEST", 0),
                    ("TEXTAFTER", 0),
                    ("TEXTBEFORE", 0),
                    ("TEXTJOIN", 0),
                    ("UNICHAR", 0),
                    ("UNICODE", 0),
                    ("VALUETOTEXT", 0),
                    ("VAR.P", 0),
                    ("VAR.S", 0),
                    ("WEBSERVICE", 0),
                    ("WEIBULL.DIST", 0),
                    ("XMATCH", 0),
                    ("XOR", 0),
                    ("Z.TEST", 0),
                    // Dynamic functions.
                    ("ANCHORARRAY", 1),
                    ("BYCOL", 1),
                    ("BYROW", 1),
                    ("CHOOSECOLS", 1),
                    ("CHOOSEROWS", 1),
                    ("DROP", 1),
                    ("EXPAND", 1),
                    ("HSTACK", 1),
                    ("LAMBDA", 1),
                    ("MAKEARRAY", 1),
                    ("MAP", 1),
                    ("RANDARRAY", 1),
                    ("REDUCE", 1),
                    ("SCAN", 1),
                    ("SEQUENCE", 1),
                    ("SINGLE", 1),
                    ("SORTBY", 1),
                    ("SWITCH", 1),
                    ("TAKE", 1),
                    ("TEXTSPLIT", 1),
                    ("TOCOL", 1),
                    ("TOROW", 1),
                    ("UNIQUE", 1),
                    ("VSTACK", 1),
                    ("WRAPCOLS", 1),
                    ("WRAPROWS", 1),
                    ("XLOOKUP", 1),
                    // Special case dynamic functions.
                    ("FILTER", 2),
                    ("SORT", 2),
                    ("PY", 2),
                ])
            })
            .get(function)
    }
}

impl From<&str> for Formula {
    fn from(value: &str) -> Formula {
        Formula::new(value)
    }
}

impl From<&Formula> for Formula {
    fn from(value: &Formula) -> Formula {
        (*value).clone()
    }
}
