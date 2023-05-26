// format - A module for representing Excel worksheet formulas.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use regex::Regex;
use std::borrow::Cow;

/// The Formula struct is used to define a worksheet formula.
///
/// The `Formula` struct creates a formula type that can be used to write
/// worksheet formulas.
///
/// In general you would use the
/// [`worksheet.write_formula()`](crate::Worksheet::write_formula) with a string
/// representation of the formula, like this:
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
/// generic [`worksheet.write()`](crate::Worksheet::write) method:
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
/// only display the "0" result. Examples of such applications are Excel viewers,
/// PDF converters, and some mobile device applications.
///
/// If required, it is also possible to specify the calculated result of the
/// formula using the [`worksheet.set_formula_result()`] method or the
/// [`formula.set_result()`](Formula::set_result) method:
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
/// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_formula_result.png">
///
/// One common spreadsheet application where the formula recalculation doesn't
/// work is `LibreOffice` (see the following [issue report]). If you wish to
/// force recalculation in `LibreOffice` you can use the
/// [`worksheet.set_formula_result_default()`] method to set the default result
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
/// [`worksheet.set_formula_result()`]: crate::Worksheet::set_formula_result
/// [`worksheet.set_formula_result_default()`]:
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
///   "@"](#the-implicit-intersection-operator-)
/// - `ANCHORARRAY`:  Explained below in [The Spilled Range Operator
///   "#"](#the-spilled-range-operator-)
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
///  In `rust_xlsxwriter` you can use the [`worksheet.write_array_formula()`]
/// function to get a static/CSE range and
/// [`worksheet.write_dynamic_array_formula()`] or
/// [`worksheet.write_dynamic_formula()`] to get a dynamic range.
///
/// [`worksheet.write_array_formula()`]: crate::Worksheet::write_array_formula
/// [`worksheet.write_dynamic_formula()`]:
///     crate::Worksheet::write_dynamic_formula
/// [`worksheet.write_dynamic_array_formula()`]:
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
/// In Excel 365, and with [`worksheet.write_dynamic_formula()`] in
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
/// [`worksheet.write_array_formula()`] or
/// [`worksheet.write_dynamic_array_formula()`]
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
/// ## The Excel 365 LAMBDA() function
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
/// A rust xlsxwriter example that replicates the described Excel functionality
/// is shown below:
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
/// When written using [`worksheet.write_formula()`] these functions need to be
/// fully qualified with a prefix such as `_xlfn.`, as shown the table in the
/// next section below.
///
/// [`worksheet.write_formula()`]: crate::Worksheet::method.write_formula
///
/// If the prefix isn't included you will get an Excel function name error. For
/// example:
///
/// ```text
///     worksheet.write_formula(0, 0, "=STDEV.S(B1:B5)")?;
/// ```
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/working_with_formulas3.png">
///
/// If the `_xlfn.` prefix is included you will get the correct result:
///
/// ```text
///     worksheet.write_formula(0, 0, "=_xlfn.STDEV.S(B1:B5)")?;
/// ```
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/working_with_formulas2.png">
///
/// Note that the function is displayed by Excel without the prefix.
///
/// Alternatively you can use the [`worksheet.use_future_functions()`] function
/// to have `rust_xlsxwriter` automatically handle future functions for you:
///
/// [`worksheet.use_future_functions()`]: crate::Worksheet::use_future_functions
///
/// ```text
///    worksheet.use_future_functions(true);
///    worksheet.write_formula(0, 0, "=STDEV.S(B1:B5)")?;
/// ```
///
/// Or if you are using a [`Formula`] struct you can use the
/// [`Formula::use_future_functions()`](Formula::use_future_functions) method:
///
/// ```text
///     worksheet.write_formula(0, 0, Formula::new("=STDEV.S(B1:B5)").use_future_functions())?;
/// ```
///
/// This will give the same correct result as the image above.
///
///
/// ## List of Future Functions
///
/// The following list is taken from [MS XLSX extensions documentation on future
/// functions].
///
/// [MS XLSX extensions documentation on future functions]:
///     http://msdn.microsoft.com/en-us/library/dd907480%28v=office.12%29.aspx
///

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
/// | `_xlfn.IMCOSH`                   |
/// | `_xlfn.IMCOT`                    |
/// | `_xlfn.IMCSCH`                   |
/// | `_xlfn.IMCSC`                    |
/// | `_xlfn.IMSECH`                   |
/// | `_xlfn.IMSEC`                    |
/// | `_xlfn.IMSINH`                   |
/// | `_xlfn.IMTAN`                    |
/// | `_xlfn.ISFORMULA`                |
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
/// support](dynamic_arrays.md) section are also future functions, however the
/// `rust_xlsxwriter` library automatically adds the required prefixes on the
/// fly so you don't have to add them explicitly.

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
///    [`worksheet.write_array_formula()`] or
///    [`worksheet.write_dynamic_array_formula()`] (see also [Dynamic Array
///    support](dynamic_arrays.md)).
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
pub struct Formula {
    formula_string: String,
    expand_future_functions: bool,
    expand_table_functions: bool,
    pub(crate) result: Box<str>,
}

impl Formula {
    /// Create a new `Formula` struct instance.
    pub fn new(formula: impl Into<String>) -> Formula {
        Formula {
            formula_string: formula.into(),
            expand_future_functions: false,
            expand_table_functions: false,
            result: Box::from(""),
        }
    }

    /// Specify the result of a formula.
    ///
    /// As explained above in the section on [Formula
    /// Results](#formula-results) it is occasionally necessary to specify the
    /// result of a formula. This can be done using the `set_result()` method.
    ///
    /// # Arguments
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

    /// Enable the use of newer Excel future functions in the formula.
    ///
    /// As explained above in [Formulas added in Excel 2010 and
    /// later](#formulas-added-in-excel-2010-and-later), functions have been
    /// added to Excel which weren't defined in the original file specification.
    /// These functions are referred to by Microsoft as "Future Functions".
    ///
    /// When written using
    /// [`write_formula()`](crate::Worksheet::write_formula()) these functions
    /// need to be fully qualified with a prefix such as `_xlfn.`
    ///
    /// Alternatively you can use the
    /// [`worksheet.use_future_functions()`](crate::Worksheet::use_future_functions)
    /// function to have `rust_xlsxwriter` automatically handle future functions
    /// for you, or use a [`Formula`] struct and the
    /// [`Formula::use_future_functions()`](Formula::use_future_functions)
    /// method, see below.
    ///
    /// # Examples
    ///
    /// The following example demonstrates different ways to handle writing
    /// Future Functions to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_use_future_functions.rs
    /// #
    /// # use rust_xlsxwriter::{Formula, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // The following is a "Future" function and will generate a "#NAME?" warning
    ///     // in Excel.
    ///     worksheet.write_formula(0, 0, "=ISFORMULA($B$1)")?;
    ///
    ///     // The following adds the required prefix. This will work without a warning.
    ///     worksheet.write_formula(1, 0, "=_xlfn.ISFORMULA($B$1)")?;
    ///
    ///     // The following uses a Formula object and expands out any future functions.
    ///     // This also works without a warning.
    ///     worksheet.write_formula(2, 0, Formula::new("=ISFORMULA($B$1)").use_future_functions())?;
    ///
    ///     // The following expands out all future functions used in the worksheet from
    ///     // this point forward. This also works without a warning.
    ///     worksheet.use_future_functions(true);
    ///     worksheet.write_formula(3, 0, "=ISFORMULA($B$1)")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_use_future_functions.png">
    ///
    pub fn use_future_functions(mut self) -> Formula {
        self.expand_future_functions = true;
        self
    }

    /// TODO
    pub fn use_table_functions(mut self) -> Formula {
        self.expand_table_functions = true;
        self
    }

    // Check of a dynamic function/formula.
    pub(crate) fn is_dynamic_function(&self) -> bool {
        lazy_static! {
            static ref DYNAMIC_FUNCTION: Regex = Regex::new(
                r"\b(ANCHORARRAY|BYCOL|BYROW|CHOOSECOLS|CHOOSEROWS|DROP|EXPAND|FILTER|HSTACK|LAMBDA|MAKEARRAY|MAP|RANDARRAY|REDUCE|SCAN|SEQUENCE|SINGLE|SORT|SORTBY|SWITCH|TAKE|TEXTSPLIT|TOCOL|TOROW|UNIQUE|VSTACK|WRAPCOLS|WRAPROWS|XLOOKUP)\("
            )
            .unwrap();
        }
        DYNAMIC_FUNCTION.is_match(&self.formula_string)
    }

    // Utility method to optionally strip equal sign and array braces from a
    // formula and also expand out future and dynamic array formulas.
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

        let formula = if self.expand_table_functions {
            Self::escape_table_functions(&formula)
        } else {
            formula
        };

        Box::from(formula)
    }

    // Escape/expand the dynamic formula _xlfn functions.
    fn escape_dynamic_formulas1(formula: &str) -> Cow<str> {
        lazy_static! {
            static ref XLFN: Regex = Regex::new(
                r"\b(ANCHORARRAY|BYCOL|BYROW|CHOOSECOLS|CHOOSEROWS|DROP|EXPAND|HSTACK|LAMBDA|MAKEARRAY|MAP|RANDARRAY|REDUCE|SCAN|SEQUENCE|SINGLE|SORTBY|SWITCH|TAKE|TEXTSPLIT|TOCOL|TOROW|UNIQUE|VSTACK|WRAPCOLS|WRAPROWS|XLOOKUP)\("
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
                r"\b(ACOTH|ACOT|AGGREGATE|ARABIC|ARRAYTOTEXT|BASE|BETA.DIST|BETA.INV|BINOM.DIST.RANGE|BINOM.DIST|BINOM.INV|BITAND|BITLSHIFT|BITOR|BITRSHIFT|BITXOR|CEILING.MATH|CEILING.PRECISE|CHISQ.DIST.RT|CHISQ.DIST|CHISQ.INV.RT|CHISQ.INV|CHISQ.TEST|COMBINA|CONCAT|CONFIDENCE.NORM|CONFIDENCE.T|COTH|COT|COVARIANCE.P|COVARIANCE.S|CSCH|CSC|DAYS|DECIMAL|ERF.PRECISE|ERFC.PRECISE|EXPON.DIST|F.DIST.RT|F.DIST|F.INV.RT|F.INV|F.TEST|FILTERXML|FLOOR.MATH|FLOOR.PRECISE|FORECAST.ETS.CONFINT|FORECAST.ETS.SEASONALITY|FORECAST.ETS.STAT|FORECAST.ETS|FORECAST.LINEAR|FORMULATEXT|GAMMA.DIST|GAMMA.INV|GAMMALN.PRECISE|GAMMA|GAUSS|HYPGEOM.DIST|IFNA|IFS|IMCOSH|IMCOT|IMCSCH|IMCSC|IMSECH|IMSEC|IMSINH|IMTAN|ISFORMULA|ISOMITTED|ISOWEEKNUM|LET|LOGNORM.DIST|LOGNORM.INV|MAXIFS|MINIFS|MODE.MULT|MODE.SNGL|MUNIT|NEGBINOM.DIST|NORM.DIST|NORM.INV|NORM.S.DIST|NORM.S.INV|NUMBERVALUE|PDURATION|PERCENTILE.EXC|PERCENTILE.INC|PERCENTRANK.EXC|PERCENTRANK.INC|PERMUTATIONA|PHI|POISSON.DIST|QUARTILE.EXC|QUARTILE.INC|QUERYSTRING|RANK.AVG|RANK.EQ|RRI|SECH|SEC|SHEETS|SHEET|SKEW.P|STDEV.P|STDEV.S|T.DIST.2T|T.DIST.RT|T.DIST|T.INV.2T|T.INV|T.TEST|TEXTAFTER|TEXTBEFORE|TEXTJOIN|UNICHAR|UNICODE|VALUETOTEXT|VAR.P|VAR.S|WEBSERVICE|WEIBULL.DIST|XMATCH|XOR|Z.TEST)\("
            )
            .unwrap();
        }
        FUTURE.replace_all(formula, "_xlfn.$1(")
    }

    // Escape/expand table functions.
    fn escape_table_functions(formula: &str) -> Cow<str> {
        // Convert Excel 2010 "@" table ref to 2007 "#This Row".
        lazy_static! {
            static ref TABLE: Regex = Regex::new(r"@").unwrap();
        }
        TABLE.replace_all(formula, "[#This Row],")
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

        for &(formula_string, expected) in &formula_strings {
            let formula = Formula::new(formula_string);
            let prepared_formula = formula.expand_formula(true);
            assert_eq!(prepared_formula.as_ref(), expected);

            let formula = Formula::new(formula_string).use_future_functions();
            let prepared_formula = formula.expand_formula(false);
            assert_eq!(prepared_formula.as_ref(), expected);
        }
    }
}
