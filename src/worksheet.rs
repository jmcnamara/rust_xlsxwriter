// worksheet - A module for creating the Excel sheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
use chrono::{Datelike, NaiveDate, NaiveDateTime, NaiveTime};
use regex::Regex;
use std::borrow::Cow;
use std::cmp;
use std::collections::HashMap;
use std::io::Write;
use std::mem;

use crate::error::XlsxError;
use crate::format::Format;
use crate::shared_strings_table::SharedStringsTable;
use crate::utility;
use crate::xmlwriter::XMLWriter;

/// Integer type to represent a zero indexed row number. Excel's limit for rows
/// in a worksheet is 1,048,576.
pub type RowNum = u32;

/// Integer type to represent a zero indexed column number. Excel's limit for
/// columns in a worksheet is 16,384.
pub type ColNum = u16;

const ROW_MAX: RowNum = 1_048_576;
const COL_MAX: ColNum = 16_384;
const MAX_STRING_LEN: u16 = 32_767;
const DEFAULT_ROW_HEIGHT: f64 = 15.0;
const DEFAULT_COL_WIDTH: f64 = 8.43;

/// The worksheet struct represents an Excel worksheet. It handles operations
/// such as writing data to cells or formatting the worksheet layout.
///
/// <img src="https://rustxlsxwriter.github.io/images/demo.png">
///
/// # Examples
///
/// Sample code to generate the Excel file shown above.
///
/// ```rust
/// # // This code is available in examples/app_demo.rs
/// #
/// use chrono::NaiveDate;
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("demo.xlsx");
///
///     // Create some formats to use in the worksheet.
///     let bold_format = Format::new().set_bold();
///     let decimal_format = Format::new().set_num_format("0.000");
///     let date_format = Format::new().set_num_format("yyyy-mm-dd");
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Set the column width for clarity.
///     worksheet.set_column_width(0, 15)?;
///
///     // Write a string without formatting.
///     worksheet.write_string_only(0, 0, "Hello")?;
///
///     // Write a string with the bold format defined above.
///     worksheet.write_string(1, 0, "World", &bold_format)?;
///
///     // Write some numbers.
///     worksheet.write_number_only(2, 0, 1)?;
///     worksheet.write_number_only(3, 0, 2.34)?;
///
///     // Write a number with formatting.
///     worksheet.write_number(4, 0, 3.00, &decimal_format)?;
///
///     // Write a formula.
///     worksheet.write_formula_only(5, 0, "=SIN(PI()/4)")?;
///
///     // Write the date .
///     let date = NaiveDate::from_ymd(2023, 1, 25);
///     worksheet.write_date(6, 0, date, &date_format)?;
///
///     workbook.close()?;
///
///     Ok(())
/// }
/// ```
pub struct Worksheet {
    pub(crate) writer: XMLWriter,
    pub(crate) name: String,
    pub(crate) selected: bool,
    pub(crate) uses_string_table: bool,
    pub(crate) has_dynamic_arrays: bool,
    table: HashMap<RowNum, HashMap<ColNum, CellType>>,
    col_names: HashMap<ColNum, String>,
    dimensions: WorksheetDimensions,
    pub(crate) xf_formats: Vec<Format>,
    xf_indices: HashMap<String, u32>,
    global_xf_indices: Vec<u32>,
    changed_rows: HashMap<RowNum, RowOptions>,
    changed_cols: HashMap<ColNum, ColOptions>,
    page_setup_changed: bool,
    paper_size: u8,
    default_page_order: bool,
    right_to_left: bool,
    portrait: bool,
    page_view: PageView,
    zoom: u16,
    header: String,
    footer: String,
    head_footer_changed: bool,
    margin_left: f64,
    margin_right: f64,
    margin_top: f64,
    margin_bottom: f64,
    margin_header: f64,
    margin_footer: f64,
    default_result: String,
    use_future_functions: bool,
}

impl Worksheet {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Worksheet struct.
    pub(crate) fn new(name: String) -> Worksheet {
        let writer = XMLWriter::new();
        let table: HashMap<RowNum, HashMap<ColNum, CellType>> = HashMap::new();
        let col_names: HashMap<ColNum, String> = HashMap::new();
        let changed_rows: HashMap<RowNum, RowOptions> = HashMap::new();
        let changed_cols: HashMap<ColNum, ColOptions> = HashMap::new();
        let default_format = Format::new();
        let xf_indices = HashMap::from([(default_format.format_key(), 0)]);

        // Initialize the min and max dimensions with their opposite value.
        let dimensions = WorksheetDimensions {
            row_min: ROW_MAX,
            col_min: COL_MAX,
            row_max: 0,
            col_max: 0,
        };

        Worksheet {
            writer,
            name,
            selected: false,
            uses_string_table: false,
            has_dynamic_arrays: false,
            table,
            col_names,
            dimensions,
            xf_formats: vec![default_format],
            xf_indices,
            global_xf_indices: vec![],
            changed_rows,
            changed_cols,
            page_setup_changed: false,
            paper_size: 0,
            default_page_order: true,
            right_to_left: false,
            portrait: true,
            page_view: PageView::Normal,
            zoom: 100,
            header: "".to_string(),
            footer: "".to_string(),
            head_footer_changed: false,
            margin_left: 0.7,
            margin_right: 0.7,
            margin_top: 0.75,
            margin_bottom: 0.75,
            margin_header: 0.3,
            margin_footer: 0.3,
            default_result: "0".to_string(),
            use_future_functions: false,
        }
    }

    /// Set the worksheet name.
    ///
    /// Set the worksheet name. If no name is set the default Excel convention
    /// will be followed (Sheet1, Sheet2, etc.) in the order the worksheets are
    /// created.
    ///
    /// # Arguments
    ///
    /// * `name` - The worksheet name. It must follow the Excel rules, shown
    ///   below.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name cannot be
    ///   blank.
    /// * [`XlsxError::SheetnameLengthExceeded`] - Worksheet name exceeds
    ///   Excel's limit of 31 characters.
    /// * [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name
    ///   cannot contain invalid characters: `[ ] : * ? / \`
    /// * [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name
    ///   cannot start or end with an apostrophe.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting user defined worksheet names
    /// and the default values when a name isn't set.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new("worksheets.xlsx");
    ///
    ///     _ = workbook.add_worksheet();                     // Sheet1
    ///     _ = workbook.add_worksheet().set_name("Foglio2"); // Foglio2
    ///     _ = workbook.add_worksheet().set_name("Data");    // Data
    ///     _ = workbook.add_worksheet();                     // Sheet4
    ///
    /// #    workbook.close()?;
    /// #
    /// #    Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_name.png">
    ///
    /// The worksheet name must be a valid Excel worksheet name, i.e:
    ///
    /// * The name is less than 32 characters.
    /// * The name isn't blank.
    /// * The name doesn't contain any of the characters: `[ ] : * ? / \`.
    /// * The name doesn't start or end with an apostrophe.
    /// * The name shouldn't be "History" (case-insensitive) since that is
    ///   reserved by Excel.
    /// * It must not be a duplicate of another worksheet name used in the
    ///   workbook.
    ///
    /// The rules for worksheet names in Excel are explained in the [Microsoft
    /// Office documentation].
    ///
    /// [Microsoft Office documentation]:
    ///     https://support.office.com/en-ie/article/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
    ///
    pub fn set_name(&mut self, name: &str) -> Result<&mut Worksheet, XlsxError> {
        // Check that the sheet name isn't blank.
        if name.is_empty() {
            return Err(XlsxError::SheetnameCannotBeBlank);
        }

        // Check that sheet sheetname is <= 31, an Excel limit.
        if name.chars().count() > 31 {
            return Err(XlsxError::SheetnameLengthExceeded(name.to_string()));
        }

        // Check that sheetname doesn't contain any invalid characters.
        let re = Regex::new(r"[\[\]:*?/\\]").unwrap();
        if re.is_match(name) {
            return Err(XlsxError::SheetnameContainsInvalidCharacter(
                name.to_string(),
            ));
        }

        // Check that sheetname doesn't start or end with an apostrophe.
        if name.starts_with('\'') || name.ends_with('\'') {
            return Err(XlsxError::SheetnameStartsOrEndsWithApostrophe(
                name.to_string(),
            ));
        }

        self.name = name.to_string();

        Ok(self)
    }

    /// Write a formatted number to a worksheet cell.
    ///
    /// Write a number with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's [`f64`] type. This method will accept any
    /// rust type that will convert [`Into`] a f64. These include i8, u8, i16,
    /// u16, i32, u32 and f32 but not i64 or u64. IEEE 754 Doubles and f64 have
    /// around 15 digits of precision. Anything beyond that cannot be stored by
    /// Excel as a number without loss of precision and may need to be stored as
    /// a string instead.
    ///
    /// [IEEE 754]: https://en.wikipedia.org/wiki/IEEE_754
    ///
    ///  Excel doesn't have handling for NaN or INF floating point numbers.
    ///  These will be stored as the strings "Nan", "INF", and "-INF" strings
    ///  instead.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `number` - The number to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting different formatting for
    /// numbers in an Excel worksheet.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new("numbers.xlsx");
    ///
    ///     // Create some formats to use with the numbers below.
    ///     let number_format = Format::new().set_num_format("#,##0.00");
    ///     let currency_format = Format::new().set_num_format("€#,##0.00");
    ///     let percentage_format = Format::new().set_num_format("0.0%");
    ///     let bold_italic_format = Format::new().set_bold().set_italic();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_number(0, 0, 1234.5, &number_format)?;
    ///     worksheet.write_number(1, 0, 1234.5, &currency_format)?;
    ///     worksheet.write_number(2, 0, 0.3300, &percentage_format)?;
    ///     worksheet.write_number(3, 0, 1234.5, &bold_italic_format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_number.png">
    ///
    ///
    pub fn write_number<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: T,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Into<f64>,
    {
        // Store the cell data.
        self.store_number(row, col, number.into(), Some(format))
    }

    /// Write an unformatted number to a cell.
    ///
    /// Write an unformatted number to a worksheet cell. This is similar to
    /// [`write_number()`](Worksheet::write_number()) except you don' have to
    /// supply a [`Format`] so it is useful for writing raw data.
    ///
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's [`f64`] type. This method will accept any
    /// rust type that will convert [`Into`] a f64. These include i8, u8, i16,
    /// u16, i32, u32 and f32 but not i64 or u64. IEEE 754 Doubles and f64 have
    /// around 15 digits of precision. Anything beyond that cannot be stored by
    /// Excel as a number without loss of precision and may need to be stored as
    /// a string instead.
    ///
    /// [IEEE 754]: https://en.wikipedia.org/wiki/IEEE_754
    ///
    ///  Excel doesn't have handling for NaN or INF floating point numbers.
    ///  These will be stored as the strings "Nan", "INF", and "-INF" strings
    ///  instead.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `number` - The number to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing unformatted numbers to an
    /// Excel worksheet. Any numeric type that will convert [`Into`] f64 can be
    /// transferred to Excel.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new("numbers.xlsx");
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some different rust number types to a worksheet.
    ///     // Note, u64 isn't supported by Excel.
    ///     worksheet.write_number_only(0, 0, 1_u8)?;
    ///     worksheet.write_number_only(1, 0, 2_i16)?;
    ///     worksheet.write_number_only(2, 0, 3_u32)?;
    ///     worksheet.write_number_only(3, 0, 4_f32)?;
    ///     worksheet.write_number_only(4, 0, 5_f64)?;
    ///
    ///     // Write some numbers with implicit types.
    ///     worksheet.write_number_only(5, 0, 1234)?;
    ///     worksheet.write_number_only(6, 0, 1234.5)?;
    ///
    ///     // Note Excel normally ignores trailing decimal zeros
    ///     // when the number is unformatted.
    ///     worksheet.write_number_only(7, 0, 1234.50000)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_number_only.png">
    ///
    pub fn write_number_only<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: T,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Into<f64>,
    {
        // Store the cell data.
        self.store_number(row, col, number.into(), None)
    }

    /// Write a formatted string to a worksheet cell.
    ///
    /// Write a string with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// Excel only supports UTF-8 text in the xlsx file format. Any rust UTF-8
    /// encoded string can be written with this method. The maximum string
    /// size supported by Excel is 32,767 characters.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `string` - The string to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// * [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting different formatting for
    /// numbers in an Excel worksheet.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file.
    ///     let mut workbook = Workbook::new("strings.xlsx");
    ///
    ///     // Create some formats to use in the worksheet.
    ///     let bold_format = Format::new().set_bold();
    ///     let italic_format = Format::new().set_italic();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some strings with formatting.
    ///     worksheet.write_string(0, 0, "Hello",     &bold_format)?;
    ///     worksheet.write_string(1, 0, "שָׁלוֹם",      &bold_format)?;
    ///     worksheet.write_string(2, 0, "नमस्ते",      &italic_format)?;
    ///     worksheet.write_string(3, 0, "こんにちは", &italic_format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_string.png">
    ///
    pub fn write_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string, Some(format))
    }

    /// Write an unformatted string to a worksheet cell.
    ///
    /// Write an unformatted string to a worksheet cell. This is similar to
    /// [`write_string()`](Worksheet::write_string()) except you don't have to
    /// supply a [`Format`] so it is useful for writing raw data.
    ///
    /// Excel only supports UTF-8 text in the xlsx file format. Any rust UTF-8
    /// encoded string can be written with this method. The maximum string
    /// size supported by Excel is 32,767 characters.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `string` - The string to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// * [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing some strings to a worksheet. The
    /// UTF-8 strings are taken from the UTF-8 example in the [Rust Programming
    /// Language] book.
    ///
    /// [Rust Programming Language]:  https://doc.rust-lang.org/book/ch08-02-strings.html#creating-a-new-string
    ///
    /// ```
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #   // Create a new Excel file.
    /// #   let mut workbook = Workbook::new("strings.xlsx");
    /// #
    /// #   // Add a worksheet to the workbook.
    /// #   let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write some strings to the worksheet.
    ///     worksheet.write_string_only(0,  0, "السلام عليكم")?;
    ///     worksheet.write_string_only(1,  0, "Dobrý den")?;
    ///     worksheet.write_string_only(2,  0, "Hello")?;
    ///     worksheet.write_string_only(3,  0, "שָׁלוֹם")?;
    ///     worksheet.write_string_only(4,  0, "नमस्ते")?;
    ///     worksheet.write_string_only(5,  0, "こんにちは")?;
    ///     worksheet.write_string_only(6,  0, "안녕하세요")?;
    ///     worksheet.write_string_only(7,  0, "你好")?;
    ///     worksheet.write_string_only(8,  0, "Olá")?;
    ///     worksheet.write_string_only(9,  0, "Здравствуйте")?;
    ///     worksheet.write_string_only(10, 0, "Hola")?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_string_only.png">
    ///
    pub fn write_string_only(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string, None)
    }

    /// Write a formatted formula to a worksheet cell.
    ///
    /// Write a formula with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// See also the `rust_xlsxwriter` documentation on [Working with Formulas].
    ///
    /// [Working with Formulas]: https://rustxlsxwriter.github.io/formulas/intro.html
    ///
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `formula` - The formula to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formulas with formatting to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formulas.xlsx");
    /// #
    ///     // Create some formats to use in the worksheet.
    ///     let bold_format = Format::new().set_bold();
    ///     let italic_format = Format::new().set_italic();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some formulas with formatting.
    ///     worksheet.write_formula(0, 0, "=1+2+3", &bold_format)?;
    ///     worksheet.write_formula(1, 0, "=A1*2", &bold_format)?;
    ///     worksheet.write_formula(2, 0, "=SIN(PI()/4)", &italic_format)?;
    ///     worksheet.write_formula(3, 0, "=AVERAGE(1, 2, 3, 4)", &italic_format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_formula.png">
    ///
    pub fn write_formula(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_formula(row, col, formula, Some(format))
    }

    /// Write an unformatted formula to a worksheet cell.
    ///
    /// Write an unformatted formula to a worksheet cell. This is similar to
    /// [`write_formula()`](Worksheet::write_formula()) except you don't have to
    /// supply a [`Format`] so it is useful for writing raw data.
    ///
    /// See also the `rust_xlsxwriter` documentation on [Working with Formulas].
    ///
    /// [Working with Formulas]: https://rustxlsxwriter.github.io/formulas/intro.html
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `formula` - The formula to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formulas with formatting to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_formula_only.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formulas.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some formulas to the worksheet.
    ///     worksheet.write_formula_only(0, 0, "=B3 + B4")?;
    ///     worksheet.write_formula_only(1, 0, "=SIN(PI()/4)")?;
    ///     worksheet.write_formula_only(2, 0, "=SUM(B1:B5)")?;
    ///     worksheet.write_formula_only(3, 0, r#"=IF(A3>1,"Yes", "No")"#)?;
    ///     worksheet.write_formula_only(4, 0, "=AVERAGE(1, 2, 3, 4)")?;
    ///     worksheet.write_formula_only(5, 0, r#"=DATEVALUE("1-Jan-2023")"#)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_formula_only.png">
    ///
    pub fn write_formula_only(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_formula(row, col, formula, None)
    }

    /// Write a formatted array formula to a worksheet cell.
    ///
    /// Write an array formula with formatting to a worksheet cell. The format
    /// is set via a [`Format`] struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// The `write_array_formula()` method writes an array formula to a cell
    /// range. In Excel an array formula is a formula that performs a
    /// calculation on a range of values. It can return a single value or a
    /// range/"array" of values.
    ///
    /// An array formula is displayed with a pair of curly brackets around the
    /// formula like this: `{=SUM(A1:B1*A2:B2)}`. The `write_array_formula()`
    /// method doesn't require actually require these so you can omit them in
    /// the formula, and the equal sign, if you wish like this:
    /// `SUM(A1:B1*A2:B2)`.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below.
    ///
    /// # Arguments
    ///
    /// * `first_row` - The first row of the range. (All zero indexed.)
    /// * `first_col` - The first row of the range.
    /// * `last_row` - The last row of the range.
    /// * `last_col` - The last row of the range.
    /// * `formula` - The formula to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array formula with
    /// formatting to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_array_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    /// #    // Add a format.
    /// #    let bold = Format::new().set_bold();
    /// #
    /// #    // Write some test data.
    /// #    worksheet.write_number_only(0, 1, 500)?;
    /// #    worksheet.write_number_only(0, 2, 300)?;
    /// #    worksheet.write_number_only(1, 1, 10)?;
    /// #    worksheet.write_number_only(1, 2, 15)?;
    ///
    ///     // Write an array formula that returns a single value.
    ///     worksheet.write_array_formula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", &bold)?;
    ///
    /// #     // Close the file.
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_array_formula.png">
    ///
    pub fn write_array_formula(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: &str,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            formula,
            Some(format),
            false,
        )
    }

    /// Write an  array formula to a worksheet cell.
    ///
    /// The `write_array_formula_only()` method writes an array formula to a
    /// cell range. In Excel an array formula is a formula that performs a
    /// calculation on a range of values. It can return a single value or a
    /// range/"array" of values. This is similar to
    /// [`write_array_formula()`](Worksheet::write_array_formula()) except you
    /// don't have to supply a [`Format`] so it is useful for writing raw data.
    ///
    /// An array formula is displayed with a pair of curly brackets around the
    /// formula like this: `{=SUM(A1:B1*A2:B2)}`. The `write_array_formula()`
    /// method doesn't require actually require these so you can omit them in
    /// the formula, and the equal sign, if you wish like this:
    /// `SUM(A1:B1*A2:B2)`.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below.
    ///
    /// # Arguments
    ///
    /// * `first_row` - The first row of the range. (All zero indexed.)
    /// * `first_col` - The first row of the range.
    /// * `last_row` - The last row of the range.
    /// * `last_col` - The last row of the range.
    /// * `formula` - The formula to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array formulas to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_array_formula_only.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    /// #    // Write some test data.
    /// #    worksheet.write_number_only(0, 1, 500)?;
    /// #    worksheet.write_number_only(0, 2, 300)?;
    /// #    worksheet.write_number_only(1, 1, 10)?;
    /// #    worksheet.write_number_only(1, 2, 15)?;
    ///
    ///     // Write an array formula that returns a single value.
    ///     worksheet.write_array_formula_only(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}")?;
    ///
    /// #     // Close the file.
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_array_formula_only.png">
    ///
    pub fn write_array_formula_only(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: &str,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row, first_col, last_row, last_col, formula, None, false,
        )
    }

    /// Write a formatted dynamic array formula to a worksheet cell or range of
    /// cells.
    ///
    /// The `write_dynamic_array_formula()` function writes an Excel 365 dynamic
    /// array formula to a cell range. Some examples of functions that return
    /// dynamic arrays are:
    ///
    /// - `FILTER()`
    /// - `RANDARRAY()`
    /// - `SEQUENCE()`
    /// - `SORTBY()`
    /// - `SORT()`
    /// - `UNIQUE()`
    /// - `XLOOKUP()`
    /// - `XMATCH()`
    ///
    /// The format is set via a [`Format`] struct which can control the font or
    /// color or properties such as bold and italic.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below or use the
    /// [`write_dynamic_formula()`](Worksheet::write_dynamic_formula()) method.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Arguments
    ///
    /// * `first_row` - The first row of the range. (All zero indexed.)
    /// * `first_col` - The first row of the range.
    /// * `last_row` - The last row of the range.
    /// * `last_col` - The last row of the range.
    /// * `formula` - The formula to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates a static function which generally
    /// returns one value turned into a dynamic array function which returns a
    /// range of values.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_dynamic_array_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Write a dynamic formula using a static function.
    ///     worksheet.write_dynamic_array_formula(0, 1, 0, 1, "=LEN(A1:A3)", &bold)?;
    /// #
    /// #     // Write some data for the function to operate on.
    /// #     worksheet.write_string_only(0, 0, "Foo")?;
    /// #     worksheet.write_string_only(1, 0, "Food")?;
    /// #     worksheet.write_string_only(2, 0, "Frood")?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_dynamic_array_formula.png">
    ///
    pub fn write_dynamic_array_formula(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: &str,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            formula,
            Some(format),
            true,
        )
    }

    /// Write a dynamic array formula to a worksheet cell or range of cells.
    ///
    /// This method is similar to
    /// [`write_dynamic_array_formula()`](Worksheet::write_dynamic_array_formula())
    /// except that it doesn't require a [`Format`] struct.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Arguments
    ///
    /// * `first_row` - The first row of the range. (All zero indexed.)
    /// * `first_col` - The first row of the range.
    /// * `last_row` - The last row of the range.
    /// * `last_col` - The last row of the range.
    /// * `formula` - The formula to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates a static function which generally returns
    /// one value turned into a dynamic array function which returns a range of
    /// values.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_dynamic_array_formula_only.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Write a dynamic formula using a static function.
    ///     worksheet.write_dynamic_array_formula_only(0, 1, 0, 1, "=LEN(A1:A3)")?;
    /// #
    /// #     // Write some data for the function to operate on.
    /// #     worksheet.write_string_only(0, 0, "Foo")?;
    /// #     worksheet.write_string_only(1, 0, "Food")?;
    /// #     worksheet.write_string_only(2, 0, "Frood")?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_dynamic_array_formula_only.png">
    ///
    pub fn write_dynamic_array_formula_only(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: &str,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row, first_col, last_row, last_col, formula, None, true,
        )
    }

    /// Write a formatted dynamic formula to a worksheet cell.
    ///
    /// The `write_dynamic_formula()` method is similar to the
    /// [`write_dynamic_array_formula()`](Worksheet::write_dynamic_array_formula())
    /// method, shown above, except that it writes a dynamic array formula to a
    /// single cell, rather than a range. This is a syntactic shortcut since the
    /// array range isn't generally known for a dynamic range and specifying the
    /// initial cell is sufficient for Excel.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `formula` - The formula to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    pub fn write_dynamic_formula(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(row, col, row, col, formula, Some(format), true)
    }

    /// Write a dynamic formula to a worksheet cell.
    ///
    /// The `write_dynamic_formula_only()` method is similar to the
    /// [`write_dynamic_array_formula_only()`](Worksheet::write_dynamic_array_formula_only())
    /// method, shown above, except that it writes a dynamic array formula to a
    /// single cell, rather than a range. This is a syntactic shortcut since the
    /// array range isn't generally known for a dynamic range and specifying the
    /// initial cell is sufficient for Excel.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `formula` - The formula to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    pub fn write_dynamic_formula_only(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(row, col, row, col, formula, None, true)
    }

    /// Write a blank formatted worksheet cell.
    ///
    /// Write a blank cell with formatting to a worksheet cell. The format is
    /// set via a [`Format`] struct.
    ///
    /// Excel differentiates between an “Empty” cell and a “Blank” cell. An
    /// “Empty” cell is a cell which doesn’t contain data or formatting whilst a
    /// “Blank” cell doesn’t contain data but does contain formatting. Excel
    /// stores “Blank” cells but ignores “Empty” cells.
    ///
    /// The most common case for a formatted blank cell is to write a background
    /// or a border, see the example below.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a blank cell with formatting,
    /// i.e., a cell that has no data but does have formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_blank.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     let format1 = Format::new().set_background_color(XlsxColor::Yellow);
    ///
    ///     let format2 = Format::new()
    ///         .set_background_color(XlsxColor::Yellow)
    ///         .set_border(XlsxBorder::Thin);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_blank.png">
    ///
    pub fn write_blank(
        &mut self,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_blank(row, col, format)
    }

    /// Write a formatted date and time to a worksheet cell.
    ///
    /// Write a [`chrono::NaiveDateTime`] instance as an Excel datetime to a
    /// worksheet cell. The [chrono] framework provides a comprehensive range of
    /// functions and types for dealing with times and dates. The serial
    /// dates/times used by Excel don't support timezones so the `Naive` chrono
    /// variants are used.
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// [`Format`] struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// [`chrono::NaiveDateTime`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// [chrono]: https://docs.rs/chrono/latest/chrono/index.html
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `datetime` - A [`chrono::NaiveDateTime`] instance.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted datetimes in an
    /// Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_datetime.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// # use chrono::NaiveDate;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh::mm");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh::mm");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-ddThh::mm:ss");
    ///     let format4 = Format::new().set_num_format("ddd dd mmm yyyy hh::mm");
    ///     let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy hh::mm");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = NaiveDate::from_ymd(2023, 1, 25).and_hms(12, 30, 0);
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_datetime(0, 0, datetime, &format1)?;
    ///     worksheet.write_datetime(1, 0, datetime, &format2)?;
    ///     worksheet.write_datetime(2, 0, datetime, &format3)?;
    ///     worksheet.write_datetime(3, 0, datetime, &format4)?;
    ///     worksheet.write_datetime(4, 0, datetime, &format5)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_datetime.png">
    ///
    pub fn write_datetime(
        &mut self,
        row: RowNum,
        col: ColNum,
        datetime: NaiveDateTime,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = self.datetime_to_excel(datetime);

        // Store the cell data.
        self.store_number(row, col, number, Some(format))
    }

    /// Write a formatted date to a worksheet cell.
    ///
    /// Write a [`chrono::NaiveDateTime`] instance as an Excel datetime to a
    /// worksheet cell. The [chrono] framework provides a comprehensive range of
    /// functions and types for dealing with times and dates. The serial
    /// dates/times used by Excel don't support timezones so the `Naive` chrono
    /// variants are used.
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// [`Format`] struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// [`chrono::NaiveDate`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    ///
    /// [chrono]: https://docs.rs/chrono/latest/chrono/index.html
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `date` - A [`chrono::NaiveDate`] instance.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted dates in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_date.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// # use chrono::NaiveDate;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the dates below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd");
    ///     let format4 = Format::new().set_num_format("ddd dd mmm yyyy");
    ///     let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a date object.
    ///     let date = NaiveDate::from_ymd(2023, 1, 25);
    ///
    ///     // Write the date with different Excel formats.
    ///     worksheet.write_date(0, 0, date, &format1)?;
    ///     worksheet.write_date(1, 0, date, &format2)?;
    ///     worksheet.write_date(2, 0, date, &format3)?;
    ///     worksheet.write_date(3, 0, date, &format4)?;
    ///     worksheet.write_date(4, 0, date, &format5)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_date.png">
    ///
    pub fn write_date(
        &mut self,
        row: RowNum,
        col: ColNum,
        date: NaiveDate,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = self.date_to_excel(date);

        // Store the cell data.
        self.store_number(row, col, number, Some(format))
    }

    /// Write a formatted time to a worksheet cell.
    ///
    /// Write a [`chrono::NaiveDateTime`] instance as an Excel datetime to a
    /// worksheet cell. The [chrono] framework provides a comprehensive range of
    /// functions and types for dealing with times and dates. The serial
    /// dates/times used by Excel don't support timezones so the `Naive` chrono
    /// variants are used.
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// [`Format`] struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// [`chrono::NaiveTime`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    ///
    /// [chrono]: https://docs.rs/chrono/latest/chrono/index.html
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `time` - A [`chrono::NaiveTime`] instance.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted times in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_time.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// # use chrono::NaiveTime;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the times below.
    ///     let format1 = Format::new().set_num_format("h::mm");
    ///     let format2 = Format::new().set_num_format("hh::mm");
    ///     let format3 = Format::new().set_num_format("hh::mm:ss");
    ///     let format4 = Format::new().set_num_format("hh::mm:ss.000");
    ///     let format5 = Format::new().set_num_format("h::mm AM/PM");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a time object.
    ///     let time = NaiveTime::from_hms_milli(2, 59, 3, 456);
    ///
    ///     // Write the time with different Excel formats.
    ///     worksheet.write_time(0, 0, time, &format1)?;
    ///     worksheet.write_time(1, 0, time, &format2)?;
    ///     worksheet.write_time(2, 0, time, &format3)?;
    ///     worksheet.write_time(3, 0, time, &format4)?;
    ///     worksheet.write_time(4, 0, time, &format5)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_time.png">
    ///
    pub fn write_time(
        &mut self,
        row: RowNum,
        col: ColNum,
        time: NaiveTime,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = self.time_to_excel(time);

        // Store the cell data.
        self.store_number(row, col, number, Some(format))
    }

    /// Write a formatted boolean value to a worksheet cell.
    ///
    /// Write a boolean value with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `boolean` - The boolean value to write to the cell.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted boolean values to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_boolean.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_boolean(0, 0, true, &bold)?;
    ///     worksheet.write_boolean(1, 0, false, &bold)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_boolean.png">
    ///
    ///
    pub fn write_boolean(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_boolean(row, col, boolean, Some(format))
    }

    /// Write an unformatted boolean value to a cell.
    ///
    /// Write an unformatted boolean value to a worksheet cell. This is similar to
    /// [`write_boolean()`](Worksheet::write_boolean()) except you don' have to
    /// supply a [`Format`] so it is useful for writing raw data.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `boolean` - The boolean value to write to the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing boolean values to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_boolean_only.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_boolean_only(0, 0, true)?;
    ///     worksheet.write_boolean_only(1, 0, false)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_boolean_only.png">
    ///
    pub fn write_boolean_only(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_boolean(row, col, boolean, None)
    }

    /// Set the height for a row of cells.
    ///
    /// The `set_row_height()` method is used to change the default height of a
    /// row. The height is specified in character units, where the default
    /// height is 15. Excel allows height values in increments of 0.25.
    ///
    /// To specify the height in pixels use the
    /// [`set_row_height_pixels()`](Worksheet::set_row_height_pixels()) method.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `height` - The row height in character units.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the height for a row in
    /// Excel.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string_only(0, 0, "Normal")?;
    ///     worksheet.write_string_only(2, 0, "Taller")?;
    ///
    ///     // Set the row height in Excel character units.
    ///     worksheet.set_row_height(2, 30)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_row_height.png">
    ///
    pub fn set_row_height<T>(&mut self, row: RowNum, height: T) -> Result<&mut Worksheet, XlsxError>
    where
        T: Into<f64>,
    {
        // Set a suitable column range for the row dimension check/set.
        let min_col = if self.dimensions.col_min != COL_MAX {
            self.dimensions.col_min
        } else {
            0
        };

        // Check row is in the allowed range.
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing row metadata object or create a new one.
        let height = height.into();
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => row_options.height = height,
            None => {
                let row_options = RowOptions {
                    height,
                    xf_index: 0,
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    /// Set the height for a row of cells, in pixels.
    ///
    /// The `set_row_height_pixels()` method is used to change the default height of a
    /// row. The height is specified in pixels, where the default
    /// height is 20.
    ///
    /// To specify the height in Excel's character units use the
    /// [`set_row_height()`](Worksheet::set_row_height()) method.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `height` - The row height in pixels.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the height for a row in Excel.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string_only(0, 0, "Normal")?;
    ///     worksheet.write_string_only(2, 0, "Taller")?;
    ///
    ///     // Set the row height in pixels.
    ///     worksheet.set_row_height_pixels(2, 40)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_row_height.png">
    ///
    pub fn set_row_height_pixels(
        &mut self,
        row: RowNum,
        height: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        let height = 0.75 * height as f64;

        self.set_row_height(row, height)
    }

    /// Set the format for a row of cells.
    ///
    /// The `set_row_format()` method is used to change the default format of a
    /// row. Any unformatted data written to that row will then adopt that
    /// format. Formatted data written to the row will maintain its own cell
    /// format. See the example below.
    ///
    /// A future version of this library may support automatic merging of
    /// explicit cell formatting with the row formatting but that isn't
    /// currently supported.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format for a row in Excel.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add for formats.
    ///     let bold_format = Format::new().set_bold();
    ///     let red_format = Format::new().set_font_color(XlsxColor::Red);
    ///
    ///     // Set the row format.
    ///     worksheet.set_row_format(1, &red_format)?;
    ///
    ///     // Add some unformatted text that adopts the row format.
    ///     worksheet.write_string_only(1, 0, "Hello")?;
    ///
    ///     // Add some formatted text that overrides the row format.
    ///     worksheet.write_string(1, 2, "Hello", &bold_format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_row_format.png">
    ///
    pub fn set_row_format(
        &mut self,
        row: RowNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Set a suitable column range for the row dimension check/set.
        let min_col = if self.dimensions.col_min != COL_MAX {
            self.dimensions.col_min
        } else {
            0
        };

        // Check row is in the allowed range.
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_index(format);

        // Update an existing row metadata object or create a new one.
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => row_options.xf_index = xf_index,
            None => {
                let row_options = RowOptions {
                    height: DEFAULT_ROW_HEIGHT,
                    xf_index,
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    /// Set the width for a worksheet column.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// worksheet column.
    ///
    /// The ``width`` parameter sets the column width in the same units used by
    /// Excel which is: the number of characters in the default font. The
    /// default width is 8.43 in the default font of Calibri 11. The actual
    /// relationship between a string width and a column width in Excel is
    /// complex. See the [following explanation of column
    /// widths](https://support.microsoft.com/en-us/kb/214123) from the
    /// Microsoft support documentation for more details. To set the width in
    /// pixels use the
    /// [`set_column_width_pixels()`](Worksheet::set_column_width_pixels())
    /// method.
    ///
    /// There is no way to specify "AutoFit" for a column in the Excel file
    /// format. This feature is only available at runtime from within Excel. It
    /// is possible to simulate "AutoFit" in your application by tracking the
    /// maximum width of the data in the column as your write it and then
    /// adjusting the column width at the end.
    ///
    /// # Arguments
    ///
    /// * `col` - The zero indexed column number.
    /// * `width` - The row width in character units.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the width of columns in
    /// Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_width.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string_only(0, 0, "Normal")?;
    ///     worksheet.write_string_only(0, 2, "Wider")?;
    ///     worksheet.write_string_only(0, 4, "Narrower")?;
    ///
    ///     // Set the column width in Excel character units.
    ///     worksheet.set_column_width(2, 16)?;
    ///     worksheet.set_column_width(4, 4)?;
    ///     worksheet.set_column_width(5, 4)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_column_width.png">
    ///
    pub fn set_column_width<T>(
        &mut self,
        col: ColNum,
        width: T,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Into<f64>,
    {
        let width = width.into();

        // Check if column is in the allowed range without updating dimensions.
        if col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing col metadata object or create a new one.
        match self.changed_cols.get_mut(&col) {
            Some(col_options) => col_options.width = width,
            None => {
                let col_options = ColOptions { width, xf_index: 0 };
                self.changed_cols.insert(col, col_options);
            }
        }

        Ok(self)
    }

    /// Set the width for a worksheet column in pixels.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// worksheet column.
    ///
    /// To set the width in Excel character units use the
    /// [`set_column_width()`](Worksheet::set_column_width()) method.
    ///
    /// There is no way to specify "AutoFit" for a column in the Excel file
    /// format. This feature is only available at runtime from within Excel. It
    /// is possible to simulate "AutoFit" in your application by tracking the
    /// maximum width of the data in the column as your write it and then
    /// adjusting the column width at the end.
    ///
    /// # Arguments
    ///
    /// * `col` - The zero indexed column number.
    /// * `width` - The row width in pixels.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the width of columns in Excel
    /// in pixels.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_width_pixels.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string_only(0, 0, "Normal")?;
    ///     worksheet.write_string_only(0, 2, "Wider")?;
    ///     worksheet.write_string_only(0, 4, "Narrower")?;
    ///
    ///     // Set the column width in pixels.
    ///     worksheet.set_column_width_pixels(2, 117)?;
    ///     worksheet.set_column_width_pixels(4, 33)?;
    ///     worksheet.set_column_width_pixels(5, 33)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_column_width.png">
    ///
    pub fn set_column_width_pixels(
        &mut self,
        col: ColNum,
        width: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Properties for Calibri 11.
        let max_digit_width = 7.0_f64;
        let padding = 5.0_f64;
        let mut width = width as f64;

        if width < 12.0 {
            width /= max_digit_width + padding;
        } else {
            width = (width - padding) / max_digit_width
        }

        self.set_column_width(col, width)
    }

    /// Set the format for a column of cells.
    ///
    /// The `set_column_format()` method is used to change the default format of a
    /// column. Any unformatted data written to that column will then adopt that
    /// format. Formatted data written to the column will maintain its own cell
    /// format. See the example below.
    ///
    /// A future version of this library may support automatic merging of
    /// explicit cell formatting with the column formatting but that isn't
    /// currently supported.
    ///
    /// # Arguments
    ///
    /// * `col` - The zero indexed column number.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format for a column in Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add for formats.
    ///     let bold_format = Format::new().set_bold();
    ///     let red_format = Format::new().set_font_color(XlsxColor::Red);
    ///
    ///     // Set the column format.
    ///     worksheet.set_column_format(1, &red_format)?;
    ///
    ///     // Add some unformatted text that adopts the column format.
    ///     worksheet.write_string_only(0, 1, "Hello")?;
    ///
    ///     // Add some formatted text that overrides the column format.
    ///     worksheet.write_string(2, 1, "Hello", &bold_format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_column_format.png">
    ///
    pub fn set_column_format(
        &mut self,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Set a suitable row range for the dimension check/set.
        let min_row = if self.dimensions.row_min != ROW_MAX {
            self.dimensions.row_min
        } else {
            0
        };

        // Check column is in the allowed range.
        if !self.check_dimensions(min_row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_index(format);

        // Update an existing col metadata object or create a new one.
        match self.changed_cols.get_mut(&col) {
            Some(col_options) => col_options.xf_index = xf_index,
            None => {
                let col_options = ColOptions {
                    width: DEFAULT_COL_WIDTH,
                    xf_index,
                };
                self.changed_cols.insert(col, col_options);
            }
        }

        Ok(self)
    }

    /// Write a user defined result to a worksheet formula cell.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using [`write_formula()`](Worksheet::write_formula()) or
    /// [`write_formula_only()`](Worksheet::write_formula_only()). Instead it
    /// stores the value 0 as the formula result. It then sets a global flag in
    /// the XLSX file to say that all formulas and functions should be
    /// recalculated when the file is opened.
    ///
    /// This works fine with Excel and other spreadsheet applications. However,
    /// applications that don’t have a facility to calculate formulas will only
    /// display the 0 results.
    ///
    /// If required, it is possible to specify the calculated result of a
    /// formula using the `set_formula_result()` method.
    ///
    /// # Arguments
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    /// * `result` - The formula result to write to the cell.
    ///
    /// # Warnings
    ///
    /// You will get a warning if you try to set a formula result for a cell
    /// that doesn't have a formula.
    ///
    /// # Examples
    ///
    /// The following example demonstrates manually setting the result of a
    /// formula.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_formula_result.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formulas.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet
    ///         .write_formula_only(0, 0, "1+1")?
    ///         .set_formula_result(0, 0, "2");
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_formula_result(&mut self, row: RowNum, col: ColNum, result: &str) -> &mut Worksheet {
        if let Some(columns) = self.table.get_mut(&row) {
            if let Some(cell) = columns.get_mut(&col) {
                match cell {
                    CellType::Formula {
                        formula: _,
                        xf_index: _,
                        result: cell_result,
                    } => {
                        *cell_result = result.to_string();
                    }
                    CellType::ArrayFormula {
                        formula: _,
                        xf_index: _,
                        result: cell_result,
                        is_dynamic: _,
                        range: _,
                    } => {
                        *cell_result = result.to_string();
                    }
                    _ => {
                        eprintln!("Cell ({}, {}) doesn't contain a formula.", row, col);
                    }
                }
            }
        }

        self
    }

    /// Write the default formula result for worksheet formulas.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using [`write_formula()`](Worksheet::write_formula()) or
    /// [`write_formula_only()`](Worksheet::write_formula_only()). Instead it
    /// stores the value 0 as the formula result. It then sets a global flag in
    /// the XLSX file to say that all formulas and functions should be
    /// recalculated when the file is opened.
    ///
    /// However, for LibreOffice the default formula result should be set to the
    /// empty string literal `""`, via the `set_formula_result_default()`
    /// method, to force calculation of the result.
    ///
    /// # Arguments
    ///
    /// * `result` - The default formula result to write to the cell.
    ///
    /// # Examples
    ///
    /// The following example demonstrates manually setting the default result
    /// for all non-calculated formulas in a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_formula_result_default.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formulas.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.set_formula_result_default("");
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    ///
    pub fn set_formula_result_default(&mut self, result: &str) -> &mut Worksheet {
        self.default_result = result.to_string();
        self
    }

    /// Enable the use of newer Excel future functions.
    ///
    /// Enable the use of newer Excel “future” functions without having to
    /// prefix them with with `_xlfn`.
    ///
    /// Excel 2010 and later versions added functions which weren't defined in
    /// the original file specification. These functions are referred to by
    /// Microsoft as "Future Functions". Examples of these functions are `ACOT`,
    /// `CHISQ.DIST.RT` , `CONFIDENCE.NORM`, `STDEV.P`, `STDEV.S` and
    /// `WORKDAY.INTL`.
    ///
    /// When written using [`write_formula()`](Worksheet::write_formula()) these
    /// functions need to be fully qualified with a prefix such as `_xlfn.`
    ///
    /// Alternatively you can use the `worksheet.use_future_functions()`
    /// function to have `rust_xlsxwriter` automatically handle future functions
    /// for you, see below.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an Excel "Future Function"
    /// with an implicit prefix and the use_future_functions() method.
    ///
    /// ```
    /// # // This code is available in examples/doc_working_with_formulas_future3.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("future_function.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Write a future function and automatically add the required prefix.
    ///     worksheet.use_future_functions(true);
    ///     worksheet.write_formula_only(0, 0, "=STDEV.S(B1:B5)")?;
    /// #
    /// #     // Write some data for the function to operate on.
    /// #     worksheet.write_number_only(0, 1, 1.23)?;
    /// #     worksheet.write_number_only(1, 1, 1.03)?;
    /// #     worksheet.write_number_only(2, 1, 1.20)?;
    /// #     worksheet.write_number_only(3, 1, 1.15)?;
    /// #     worksheet.write_number_only(4, 1, 1.22)?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/working_with_formulas2.png">
    ///
    pub fn use_future_functions(&mut self, enable: bool) {
        self.use_future_functions = enable;
    }

    // -----------------------------------------------------------------------
    // Worksheet page setup methods.
    // -----------------------------------------------------------------------

    /// Display the worksheet cells from right to left for some versions of
    /// Excel.
    ///
    /// The `set_right_to_left()` method is used to change the default direction
    /// of the worksheet from left-to-right, with the A1 cell in the top left,
    /// to right-to-left, with the A1 cell in the top right.
    ///
    /// This is useful when creating Arabic, Hebrew or other near or far eastern
    /// worksheets that use right-to-left as the default direction.
    ///
    /// Depending on your use case, and text, you may also need to use the
    /// [`Format::set_reading_direction()`](super::Format::set_reading_direction)
    /// method to set the direction of the text within the cells.
    ///
    /// # Examples
    ///
    /// The following example demonstrates changing the default worksheet and
    /// cell text direction changed from left-to-right to right-to-left, as
    /// required by some middle eastern versions of Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_right_to_left.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    ///     // Add the cell formats.
    ///     let format_left_to_right = Format::new().set_reading_direction(1);
    ///     let format_right_to_left = Format::new().set_reading_direction(2);
    ///
    ///     // Add a worksheet in the standard left to right direction.
    ///     let worksheet1 = workbook.add_worksheet();
    ///
    ///     // Make the column wider for clarity.
    ///     worksheet1.set_column_width(0,25)?;
    ///
    ///     // Standard direction:         | A1 | B1 | C1 | ...
    ///     worksheet1.write_string_only(0, 0, "نص عربي / English text")?;
    ///     worksheet1.write_string(1, 0, "نص عربي / English text", &format_left_to_right)?;
    ///     worksheet1.write_string(2, 0, "نص عربي / English text", &format_right_to_left)?;
    ///
    ///     // Add a worksheet and change it to right to left direction.
    ///     let worksheet2 = workbook.add_worksheet();
    ///     worksheet2.set_right_to_left();
    ///
    ///     // Make the column wider for clarity.
    ///     worksheet2.set_column_width(0, 25)?;
    ///
    ///     // Right to left direction:    ... | C1 | B1 | A1 |
    ///     worksheet2.write_string_only(0, 0, "نص عربي / English text")?;
    ///     worksheet2.write_string(1, 0, "نص عربي / English text", &format_left_to_right)?;
    ///     worksheet2.write_string(2, 0, "نص عربي / English text", &format_right_to_left)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_right_to_left.png">
    ///
    pub fn set_right_to_left(&mut self) -> &mut Worksheet {
        self.right_to_left = true;
        self
    }

    /// Set the paper type/size when printing.
    ///
    /// This method is used to set the paper format for the printed output of a
    /// worksheet. The following paper styles are available:
    ///
    /// | Index    | Paper format            | Paper size           |
    /// | :------- | :---------------------- | :------------------- |
    /// | 0        | Printer default         | Printer default      |
    /// | 1        | Letter                  | 8 1/2 x 11 in        |
    /// | 2        | Letter Small            | 8 1/2 x 11 in        |
    /// | 3        | Tabloid                 | 11 x 17 in           |
    /// | 4        | Ledger                  | 17 x 11 in           |
    /// | 5        | Legal                   | 8 1/2 x 14 in        |
    /// | 6        | Statement               | 5 1/2 x 8 1/2 in     |
    /// | 7        | Executive               | 7 1/4 x 10 1/2 in    |
    /// | 8        | A3                      | 297 x 420 mm         |
    /// | 9        | A4                      | 210 x 297 mm         |
    /// | 10       | A4 Small                | 210 x 297 mm         |
    /// | 11       | A5                      | 148 x 210 mm         |
    /// | 12       | B4                      | 250 x 354 mm         |
    /// | 13       | B5                      | 182 x 257 mm         |
    /// | 14       | Folio                   | 8 1/2 x 13 in        |
    /// | 15       | Quarto                  | 215 x 275 mm         |
    /// | 16       | ---                     | 10x14 in             |
    /// | 17       | ---                     | 11x17 in             |
    /// | 18       | Note                    | 8 1/2 x 11 in        |
    /// | 19       | Envelope 9              | 3 7/8 x 8 7/8        |
    /// | 20       | Envelope 10             | 4 1/8 x 9 1/2        |
    /// | 21       | Envelope 11             | 4 1/2 x 10 3/8       |
    /// | 22       | Envelope 12             | 4 3/4 x 11           |
    /// | 23       | Envelope 14             | 5 x 11 1/2           |
    /// | 24       | C size sheet            | ---                  |
    /// | 25       | D size sheet            | ---                  |
    /// | 26       | E size sheet            | ---                  |
    /// | 27       | Envelope DL             | 110 x 220 mm         |
    /// | 28       | Envelope C3             | 324 x 458 mm         |
    /// | 29       | Envelope C4             | 229 x 324 mm         |
    /// | 30       | Envelope C5             | 162 x 229 mm         |
    /// | 31       | Envelope C6             | 114 x 162 mm         |
    /// | 32       | Envelope C65            | 114 x 229 mm         |
    /// | 33       | Envelope B4             | 250 x 353 mm         |
    /// | 34       | Envelope B5             | 176 x 250 mm         |
    /// | 35       | Envelope B6             | 176 x 125 mm         |
    /// | 36       | Envelope                | 110 x 230 mm         |
    /// | 37       | Monarch                 | 3.875 x 7.5 in       |
    /// | 38       | Envelope                | 3 5/8 x 6 1/2 in     |
    /// | 39       | Fanfold                 | 14 7/8 x 11 in       |
    /// | 40       | German Std Fanfold      | 8 1/2 x 12 in        |
    /// | 41       | German Legal Fanfold    | 8 1/2 x 13 in        |
    ///
    /// Note, it is likely that not all of these paper types will be available
    /// to the end user since it will depend on the paper formats that the
    /// user's printer supports. Therefore, it is best to stick to standard
    /// paper types of 1 for US Letter and 9 for A4.
    ///
    /// If you do not specify a paper type the worksheet will print using the
    /// printer's default paper style.
    ///
    /// # Arguments
    ///
    /// * `paper_size` - The paper size index from the list above .
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet paper size/type for
    /// the printed output.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_paper.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     // Set the printer paper size.
    ///     worksheet.set_paper_size(9); // A4 paper size.
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn set_paper_size(&mut self, paper_size: u8) -> &mut Worksheet {
        self.paper_size = paper_size;
        self.page_setup_changed = true;
        self
    }

    /// Set the order in which pages are printed.
    ///
    /// The `set_page_order()` method is used to change the default print
    /// direction. This is referred to by Excel as the sheet "page order":
    ///
    /// The default page order is shown below for a worksheet that extends over
    /// 4 pages. The order is called "down then over":
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_page_order.png">
    ///
    /// However, by using the `set_page_order()` method the print order will be
    /// changed to "over then down".
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet printed page order.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_page_order.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set the page print to "over then down"
    ///     worksheet.set_page_order();
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn set_page_order(&mut self) -> &mut Worksheet {
        self.default_page_order = false;
        self.page_setup_changed = true;
        self
    }

    /// Set the page orientation to landscape.
    ///
    /// The `set_landscape()` method is used to set the orientation of a
    /// worksheet's printed page to landscape.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet page orientation to
    /// landscape.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_landscape.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.set_landscape();
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_landscape(&mut self) -> &mut Worksheet {
        self.portrait = false;
        self.page_setup_changed = true;
        self
    }

    /// Set the page orientation to portrait.
    ///
    ///  This `set_portrait()` method  is used to set the orientation of a
    ///  worksheet's printed page to portrait. The default worksheet orientation
    ///  is portrait, so this function is rarely required.
    ///
    pub fn set_portrait(&mut self) -> &mut Worksheet {
        self.portrait = true;
        self.page_setup_changed = true;
        self
    }

    /// Set the page view mode to normal layout.
    ///
    /// This method is used to display the worksheet in “View -> Normal”
    /// mode. This is the default.
    ///
    pub fn set_view_normal(&mut self) -> &mut Worksheet {
        self.page_view = PageView::Normal;
        self
    }

    /// Set the page view mode to page layout.
    ///
    /// This method is used to display the worksheet in “View -> Page Layout”
    /// mode.
    ///
    pub fn set_view_page_layout(&mut self) -> &mut Worksheet {
        self.page_view = PageView::PageLayout;
        self.page_setup_changed = true;
        self
    }

    /// Set the page view mode to page break preview.
    ///
    /// This method is used to display the worksheet in “View -> Page Break
    /// Preview” mode.
    ///
    pub fn set_view_page_break_preview(&mut self) -> &mut Worksheet {
        self.page_view = PageView::PageBreaks;
        self.page_setup_changed = true;
        self
    }

    /// Set the worksheet zoom factor.
    ///
    /// Set the worksheet zoom factor in the range 10 <= zoom <= 400.
    ///
    /// # Arguments
    ///
    /// * `zoom` - The worksheet zoom level.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet zoom level.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_zoom.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_string_only(0, 0, "Hello")?;
    ///     worksheet.set_zoom(200);
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_zoom.png">
    ///
    pub fn set_zoom(&mut self, zoom: u16) -> &mut Worksheet {
        if !(10..=400).contains(&zoom) {
            eprintln!("Zoom factor {} outside range: 10 <= zoom <= 400.", zoom);
            return self;
        }

        self.zoom = zoom;
        self
    }

    /// Set the printed page header caption.
    ///
    /// The `set_header()` method can be used to set the header for a worksheet.
    ///
    /// Headers and footers are generated using a string which is a combination
    /// of plain text and optional control characters.
    ///
    /// The available control characters are:
    ///
    /// | Control              | Category      | Description           |
    /// | -------------------- | ------------- | --------------------- |
    /// | `&L`                 | Alignment     | Left                  |
    /// | `&C`                 |               | Center                |
    /// | `&R`                 |               | Right                 |
    /// | `&[Page]`  or `&P`   | Information   | Page number           |
    /// | `&[Pages]` or `&N`   |               | Total number of pages |
    /// | `&[Date]`  or `&D`   |               | Date                  |
    /// | `&[Time]`  or `&T`   |               | Time                  |
    /// | `&[File]`  or `&F`   |               | File name             |
    /// | `&[Tab]`   or `&A`   |               | Worksheet name        |
    /// | `&[Path]`  or `&Z`   |               | Workbook path         |
    /// | `&fontsize`          | Font          | Font size             |
    /// | `&"font,style"`      |               | Font name and style   |
    /// | `&U`                 |               | Single underline      |
    /// | `&E`                 |               | Double underline      |
    /// | `&S`                 |               | Strikethrough         |
    /// | `&X`                 |               | Superscript           |
    /// | `&Y`                 |               | Subscript             |
    /// | `&&`                 | Miscellaneous | Literal ampersand &   |
    ///
    /// Some of the placeholder variables have a long version like `&[Page]` and
    /// a short version like `&P`. The longer version is displayed in the Excel
    /// interface but the shorter version is the way that it is stored in the
    /// file format. Either version is okay since `rust_xlsxwriter` will
    /// translate as required.
    ///
    /// Headers and footers have 3 edit areas to the left, center and right.
    /// Text can be aligned to these areas by prefixing the text with the
    /// control characters `&L`, `&C` and `&R`.
    ///
    /// For example:
    ///
    /// ```text
    /// worksheet.set_header("&LHello");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// | Hello                                                         |
    /// |                                                               |
    ///
    /// worksheet.set_header("&CHello");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                          Hello                                |
    /// |                                                               |
    ///
    /// worksheet.set_header("&RHello");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                                                         Hello |
    /// |                                                               |
    /// ```
    ///
    /// You can also have text in each of the alignment areas:
    ///
    /// ```text
    /// worksheet.set_header("&LCiao&CBello&RCielo");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// | Ciao                     Bello                          Cielo |
    /// |                                                               |
    /// ```
    ///
    /// The information control characters act as variables/templates that Excel
    /// will update/expand as the workbook or worksheet changes.
    ///
    /// ```text
    /// worksheet.set_header("&CPage &[Page] of &[Pages]");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                        Page 1 of 6                            |
    /// |                                                               |
    /// ```
    ///
    /// Times and dates are in the user's default format:
    ///
    /// ```text
    /// worksheet.set_header("&CUpdated at &[Time]");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                    Updated at 12:30 PM                        |
    /// |                                                               |
    /// ```
    ///
    /// To include a single literal ampersand `&` in a header or footer you
    /// should use a double ampersand `&&`:
    ///
    /// ```text
    /// worksheet.set_header("&CCuriouser && Curiouser - Attorneys at Law");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                   Curiouser & Curiouser                       |
    /// |                                                               |
    /// ```
    ///
    /// You can specify the font size of a section of the text by prefixing it
    /// with the control character `&n` where `n` is the font size:
    ///
    /// ```text
    /// worksheet1.set_header("&C&30Hello Big");
    /// worksheet2.set_header("&C&10Hello Small");
    /// ```
    ///
    /// You can specify the font of a section of the text by prefixing it with
    /// the control sequence `&"font,style"` where `fontname` is a font name
    /// such as Windows font descriptions: "Regular", "Italic", "Bold" or "Bold
    /// Italic": "Courier New" or "Times New Roman" and `style` is one of the
    /// standard Windows font descriptions like “Regular”, “Italic”, “Bold” or
    /// “Bold Italic”:
    ///
    /// ```text
    /// worksheet1.set_header(r#"&C&"Courier New,Italic"Hello"#);
    /// worksheet2.set_header(r#"&C&"Courier New,Bold Italic"Hello"#);
    /// worksheet3.set_header(r#"&C&"Times New Roman,Regular"Hello"#);
    /// ```
    ///
    /// It is possible to combine all of these features together to create
    /// complex headers and footers. If you set up a complex header in Excel you
    /// can transfer it to `rust_xlsxwriter` by inspecting the string in the
    /// Excel file. For example the following shows how unzip and grep the Excel
    /// XML sub-files on a Linux system. The example uses libxml's xmllint to
    /// format the XML for clarity:
    ///
    /// ```text
    /// $ unzip myfile.xlsm -d myfile
    /// $ xmllint --format `find myfile -name "*.xml" | xargs` | \
    ///     egrep "Header|Footer" | sed 's/&amp;/\&/g'
    ///
    ///  <headerFooter scaleWithDoc="0">
    ///    <oddHeader>&L&P</oddHeader>
    ///  </headerFooter>
    /// ```
    ///
    /// Note: Excel requires that the header or footer string be less than 256
    /// characters, including the control characters. Strings longer than this
    /// will not be written, and a warning will be output.
    ///
    /// # Arguments
    ///
    /// * `header` - The header string with optional control characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet header.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_header.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.set_header("&CPage &P of &N");
    ///
    /// #     worksheet.write_string_only(0, 0, "Hello")?;
    /// #     worksheet.write_string_only(200, 0, "Hello")?;
    /// #     worksheet.set_view_page_layout();
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_header.png">
    ///
    pub fn set_header(&mut self, header: &str) -> &mut Worksheet {
        let header = header
            .replace("&[Tab]", "&A")
            .replace("&[Date]", "&D")
            .replace("&[File]", "&F")
            .replace("&[Page]", "&P")
            .replace("&[Path]", "&Z")
            .replace("&[Time]", "&T")
            .replace("&[Pages]", "&N")
            .replace("&[Picture]", "&G");

        if header.chars().count() > 255 {
            eprintln!("Header string exceeds Excel's limit of 255 characters.");
            return self;
        }

        self.header = header;
        self.page_setup_changed = true;
        self.head_footer_changed = true;
        self
    }

    /// Set the printed page footer caption.
    ///
    /// The `set_footer()` method can be used to set the footer for a worksheet.
    ///
    /// See the documentation for [`set_header()`](Worksheet::set_header()) for
    /// more details on the syntax of the header/footer string.
    ///
    /// # Arguments
    ///
    /// * `footer` - The footer string with optional control characters.
    ///
    pub fn set_footer(&mut self, footer: &str) -> &mut Worksheet {
        let footer = footer
            .replace("&[Tab]", "&A")
            .replace("&[Date]", "&D")
            .replace("&[File]", "&F")
            .replace("&[Page]", "&P")
            .replace("&[Path]", "&Z")
            .replace("&[Time]", "&T")
            .replace("&[Pages]", "&N")
            .replace("&[Picture]", "&G");

        if footer.chars().count() > 255 {
            eprintln!("Footer string exceeds Excel's limit of 255 characters.");
            return self;
        }

        self.footer = footer;
        self.page_setup_changed = true;
        self.head_footer_changed = true;
        self
    }

    /// Set the page margins.
    ///
    /// The `set_margins()` method is used to set the margins of the worksheet
    /// when it is printed. The units are in inches. Specifying `-1.0` for any
    /// parameter will give the default Excel value. The defaults are shown
    /// below.
    ///
    /// # Arguments
    ///
    /// * `left` - Left margin in inches. Excel default is 0.7.
    /// * `right` - Right margin in inches. Excel default is 0.7.
    /// * `top` - Top margin in inches. Excel default is 0.75.
    /// * `bottom` - Bottom margin in inches. Excel default is 0.75.
    /// * `header` - Header margin in inches. Excel default is 0.3.
    /// * `footer` - Footer margin in inches. Excel default is 0.3.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet margins.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_margins.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new("worksheet.xlsx");
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.set_margins(1.0, 1.25, 1.5, 1.75, 0.75, 0.25);
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_margins.png">
    ///
    pub fn set_margins(
        &mut self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> &mut Worksheet {
        if left >= 0.0 {
            self.margin_left = left;
            self.page_setup_changed = true;
        }
        if right >= 0.0 {
            self.margin_right = right;
            self.page_setup_changed = true;
        }
        if top >= 0.0 {
            self.margin_top = top;
            self.page_setup_changed = true;
        }
        if bottom >= 0.0 {
            self.margin_bottom = bottom;
            self.page_setup_changed = true;
        }
        if header >= 0.0 {
            self.margin_header = header;
            self.page_setup_changed = true;
        }
        if footer >= 0.0 {
            self.margin_footer = footer;
            self.page_setup_changed = true;
        }

        self
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    // Store a number cell in the worksheet data table structure.
    fn store_number(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: f64,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Excel doesn't have a NAN type/value so write a string instead.
        if number.is_nan() {
            return self.store_string(row, col, "#NUM!", None);
        }

        // Excel doesn't have an Infinity type/value so write a string instead.
        if number.is_infinite() {
            self.store_string(row, col, "#DIV/0", None)?;
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_index(format),
            None => 0,
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Number { number, xf_index };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store a string cell in the worksheet data table structure.
    fn store_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        //  Check that the string is < Excel limit of 32767 chars.
        if string.chars().count() as u16 > MAX_STRING_LEN {
            return Err(XlsxError::MaxStringLengthExceeded);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_index(format),
            None => 0,
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::String {
            string: string.to_string(),
            xf_index,
        };

        self.insert_cell(row, col, cell);
        self.uses_string_table = true;

        Ok(self)
    }

    // Store a formula cell in the worksheet data table structure.
    fn store_formula(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Transfer to dynamic formula handling function.
        if is_dynamic_function(formula) {
            return self.store_array_formula(row, col, row, col, formula, None, true);
        }

        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_index(format),
            None => 0,
        };

        let formula = prepare_formula(formula, self.use_future_functions);

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Formula {
            formula,
            xf_index,
            result: self.default_result.clone(),
        };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store an array formula cell in the worksheet data table structure.
    #[allow(clippy::too_many_arguments)]
    fn store_array_formula(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: &str,
        format: Option<&Format>,
        is_dynamic: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(first_row, first_col)
            || !self.check_dimensions(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        let first_row = first_row;

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_index(format),
            None => 0,
        };

        let formula = prepare_formula(formula, self.use_future_functions);

        // Create the array range reference.
        let range = utility::cell_range(first_row, first_col, last_row, last_col);

        // Check for a dynamic function in a standard static array formula.
        let mut is_dynamic = is_dynamic;
        if !is_dynamic && is_dynamic_function(&formula) {
            is_dynamic = true;
        }

        if is_dynamic {
            self.has_dynamic_arrays = true;
        }

        // Create the appropriate cell type to hold the data.
        let cell = CellType::ArrayFormula {
            formula,
            xf_index,
            result: self.default_result.clone(),
            is_dynamic,
            range,
        };

        self.insert_cell(first_row, first_col, cell);

        // Pad out the rest of the area with formatted zeroes.
        for row in first_row..=last_row {
            for col in first_col..=last_col {
                if !(row == first_row && col == first_col) {
                    match format {
                        Some(format) => self.write_number(row, col, 0, format).unwrap(),
                        None => self.write_number_only(row, col, 0).unwrap(),
                    };
                }
            }
        }

        Ok(self)
    }

    // Store a blank cell in the worksheet data table structure.
    fn store_blank(
        &mut self,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_index(format);

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Blank { xf_index };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store a boolean cell in the worksheet data table structure.
    fn store_boolean(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_index(format),
            None => 0,
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Boolean { boolean, xf_index };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Insert a cell value into the worksheet data table structure.
    fn insert_cell(&mut self, row: RowNum, col: ColNum, cell: CellType) {
        match self.table.get_mut(&row) {
            Some(columns) => {
                // The row already exists. Insert/replace column value.
                columns.insert(col, cell);
            }
            None => {
                // The row doesn't exist, create a new row with columns and insert
                // the cell value.
                let mut columns: HashMap<ColNum, CellType> = HashMap::new();
                columns.insert(col, cell);
                self.table.insert(row, columns);
            }
        }
    }

    // Check that row and col are within the allowed Excel range and store max
    // and min values for use in other methods/elements.
    fn check_dimensions(&mut self, row: RowNum, col: ColNum) -> bool {
        // Check that the row an column number are withing Excel's ranges.
        if row >= ROW_MAX {
            return false;
        }
        if col >= COL_MAX {
            return false;
        }

        // Store any changes in worksheet dimensions.
        self.dimensions.row_min = cmp::min(self.dimensions.row_min, row);
        self.dimensions.col_min = cmp::min(self.dimensions.col_min, col);
        self.dimensions.row_max = cmp::max(self.dimensions.row_max, row);
        self.dimensions.col_max = cmp::max(self.dimensions.col_max, col);

        true
    }

    // Cached/faster version of utility.col_to_name() to use in the inner loop.
    fn col_to_name(&mut self, col_num: ColNum) -> String {
        if let Some(col_name) = self.col_names.get(&col_num) {
            col_name.clone()
        } else {
            let col_name = utility::col_to_name(col_num);
            self.col_names.insert(col_num, col_name.clone());
            col_name
        }
    }

    // Store local copies of unique formats passed to the write methods. These
    // indexes will be replaced by global/workbook indices before the worksheet
    // is saved.
    fn format_index(&mut self, format: &Format) -> u32 {
        let format_key = format.format_key();

        match self.xf_indices.get_mut(&format_key) {
            Some(xf_index) => *xf_index,
            None => {
                let xf_index = self.xf_formats.len() as u32;
                self.xf_formats.push(format.clone());
                self.xf_indices.insert(format_key, xf_index);
                xf_index
            }
        }
    }

    // Set the mapping between the local format indices and the global/workbook
    // indices.
    pub(crate) fn set_global_xf_indices(&mut self, workbook_indices: &[u32]) {
        self.global_xf_indices = workbook_indices.to_owned();
    }

    // Translate the cell xf_index into a global/workbook format index. We also
    // need to make sure that an unformatted cell (xf_index == 0) takes the row
    // format (if it exists) or, failing that, the column format (if that
    // exists).
    fn get_cell_xf_index(
        &mut self,
        xf_index: &u32,
        row_options: Option<&RowOptions>,
        col_num: ColNum,
    ) -> u32 {
        // The local cell format index.
        let mut xf_index = *xf_index;

        // If it is zero the cell is unformatted and we check for a row format.
        if xf_index == 0 {
            if let Some(row_options) = row_options {
                xf_index = row_options.xf_index;
            }
        }

        // If it is still zero the row was unformatted so we check for a column
        // format.
        if xf_index == 0 {
            if let Some(col_options) = self.changed_cols.get(&col_num) {
                xf_index = col_options.xf_index;
            }
        }

        // Finally convert the local format index into a global/workbook index.
        if xf_index != 0 {
            xf_index = self.global_xf_indices[xf_index as usize];
        }

        xf_index
    }

    // Notes for the date/time handling functions below.
    //
    // * Datetimes in Excel are a serial date with days counted from an epoch
    //   (generally 1899-12-31) and the time as a percentage/decimal of the
    //   milliseconds in the day.
    //
    // * Both are stored in the same f64 value, for example, 2023/01/01 12:00:00 is
    //   stored as 44927.5 with a separate numeric format like yyyy/mm/dd hh:mm.
    //
    // * Excel can also save dates in a text ISO 8601 format in "Strict Open XML
    //   Spreadsheet" format but this is rarely used in practice.
    //
    // * Excel also doesn't use timezones or try to convert or encode timezone
    //   information in any way.

    // Convert a chrono::NaiveTime to an Excel serial datetime.
    fn datetime_to_excel(&mut self, datetime: NaiveDateTime) -> f64 {
        let excel_date = self.date_to_excel(datetime.date());
        let excel_time = self.time_to_excel(datetime.time());

        excel_date + excel_time
    }

    // Convert a chrono::NaiveDate to an Excel serial date. In Excel a serial date
    // is the number of days since the epoch, which is either 1899-12-31 or
    // 1904-01-01.
    fn date_to_excel(&mut self, date: NaiveDate) -> f64 {
        let epoch = NaiveDate::from_ymd(1899, 12, 31);

        let duration = date - epoch;
        let mut excel_date = duration.num_days() as f64;

        // For legacy reasons Excel treats 1900 as a leap year. We add an additional
        // day for dates after the leapday in the 1899 epoch.
        if epoch.year() == 1899 && excel_date > 59.0 {
            excel_date += 1.0;
        }

        excel_date
    }

    // Convert a chrono::NaiveTime to an Excel time. The time portion of the Excel
    // datetime is the number of milliseconds divided by the total number of
    // milliseconds in the day.
    fn time_to_excel(&mut self, time: NaiveTime) -> f64 {
        let midnight = NaiveTime::from_hms_milli(0, 0, 0, 0);
        let duration = time - midnight;

        duration.num_milliseconds() as f64 / (24.0 * 60.0 * 60.0 * 1000.0)
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self, string_table: &mut SharedStringsTable) {
        self.writer.xml_declaration();

        // Write the worksheet element.
        self.write_worksheet();

        // Write the dimension element.

        self.write_dimension();

        // Write the sheetViews element.
        self.write_sheet_views();

        // Write the sheetFormatPr element.
        self.write_sheet_format_pr();

        // Write the cols element.
        self.write_cols();

        // Write the sheetData element.
        self.write_sheet_data(string_table);

        // Write the pageMargins element.
        self.write_page_margins();

        // Write the pageSetup element.
        if self.page_setup_changed {
            self.write_page_setup();
        }

        // Write the headerFooter element.
        if self.head_footer_changed {
            self.write_header_footer();
        }

        // Close the worksheet tag.
        self.writer.xml_end_tag("worksheet");
    }

    // Write the <worksheet> element.
    fn write_worksheet(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let xmlns_r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string();

        let attributes = vec![("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag_attr("worksheet", &attributes);
    }

    // Write the <dimension> element.
    fn write_dimension(&mut self) {
        let mut attributes = vec![];
        let mut range = "A1".to_string();

        if !self.table.is_empty() || !self.changed_rows.is_empty() || !self.changed_cols.is_empty()
        {
            range = utility::cell_range(
                self.dimensions.row_min,
                self.dimensions.col_min,
                self.dimensions.row_max,
                self.dimensions.col_max,
            );
        }

        attributes.push(("ref", range));

        self.writer.xml_empty_tag_attr("dimension", &attributes);
    }

    // Write the <sheetViews> element.
    fn write_sheet_views(&mut self) {
        self.writer.xml_start_tag("sheetViews");

        // Write the sheetView element.
        self.write_sheet_view();

        self.writer.xml_end_tag("sheetViews");
    }

    // Write the <sheetView> element.
    fn write_sheet_view(&mut self) {
        let mut attributes = vec![];

        if self.right_to_left {
            attributes.push(("rightToLeft", "1".to_string()));
        }

        if self.selected {
            attributes.push(("tabSelected", "1".to_string()));
        }

        match self.page_view {
            PageView::PageLayout => {
                attributes.push(("view", "pageLayout".to_string()));
            }
            PageView::PageBreaks => {
                attributes.push(("view", "pageBreakPreview".to_string()));
            }
            PageView::Normal => {}
        }

        if self.zoom != 100 {
            attributes.push(("zoomScale", self.zoom.to_string()));

            match self.page_view {
                PageView::PageLayout => {
                    attributes.push(("zoomScalePageLayoutView", self.zoom.to_string()));
                }
                PageView::PageBreaks => {
                    attributes.push(("zoomScaleSheetLayoutView", self.zoom.to_string()));
                }
                PageView::Normal => {
                    attributes.push(("zoomScaleNormal", self.zoom.to_string()));
                }
            }
        }

        attributes.push(("workbookViewId", "0".to_string()));

        self.writer.xml_empty_tag_attr("sheetView", &attributes);
    }

    // Write the <sheetFormatPr> element.
    fn write_sheet_format_pr(&mut self) {
        let attributes = vec![("defaultRowHeight", "15".to_string())];

        self.writer.xml_empty_tag_attr("sheetFormatPr", &attributes);
    }

    // Write the <sheetData> element.
    fn write_sheet_data(&mut self, string_table: &mut SharedStringsTable) {
        if self.table.is_empty() && self.changed_rows.is_empty() {
            self.writer.xml_empty_tag("sheetData");
        } else {
            self.writer.xml_start_tag("sheetData");
            self.write_data_table(string_table);
            self.writer.xml_end_tag("sheetData");
        }
    }

    // Write the <pageMargins> element.
    fn write_page_margins(&mut self) {
        let attributes = vec![
            ("left", self.margin_left.to_string()),
            ("right", self.margin_right.to_string()),
            ("top", self.margin_top.to_string()),
            ("bottom", self.margin_bottom.to_string()),
            ("header", self.margin_header.to_string()),
            ("footer", self.margin_footer.to_string()),
        ];

        self.writer.xml_empty_tag_attr("pageMargins", &attributes);
    }

    // Write the <pageSetup> element.
    fn write_page_setup(&mut self) {
        let mut attributes = vec![];

        if self.paper_size > 0 {
            attributes.push(("paperSize", self.paper_size.to_string()));
        }

        if !self.default_page_order {
            attributes.push(("pageOrder", "overThenDown".to_string()));
        }

        if self.portrait {
            attributes.push(("orientation", "portrait".to_string()));
        } else {
            attributes.push(("orientation", "landscape".to_string()));
        }

        attributes.push(("horizontalDpi", "200".to_string()));
        attributes.push(("verticalDpi", "200".to_string()));

        self.writer.xml_empty_tag_attr("pageSetup", &attributes);
    }

    // Write out all the row and cell data in the worksheet data table.
    fn write_data_table(&mut self, string_table: &mut SharedStringsTable) {
        let spans = self.calculate_spans();

        // Swap out the worksheet data structures so we can iterate over it and
        // still call self.write_xml() methods.
        //
        // TODO. check efficiency of this and/or alternatives.
        let mut temp_table: HashMap<RowNum, HashMap<ColNum, CellType>> = HashMap::new();
        let mut temp_changed_rows: HashMap<RowNum, RowOptions> = HashMap::new();
        mem::swap(&mut temp_table, &mut self.table);
        mem::swap(&mut temp_changed_rows, &mut self.changed_rows);

        for row_num in self.dimensions.row_min..=self.dimensions.row_max {
            let span_index = row_num / 16;
            let span = spans.get(&span_index);

            let columns = temp_table.get(&row_num);
            let row_options = temp_changed_rows.get(&row_num);

            if columns.is_some() || row_options.is_some() {
                if let Some(columns) = columns {
                    self.write_row(row_num, span, row_options, true);
                    for col_num in self.dimensions.col_min..=self.dimensions.col_max {
                        if let Some(cell) = columns.get(&col_num) {
                            match cell {
                                CellType::Number { number, xf_index } => {
                                    let xf_index =
                                        self.get_cell_xf_index(xf_index, row_options, col_num);
                                    self.write_number_cell(row_num, col_num, number, &xf_index)
                                }
                                CellType::String { string, xf_index } => {
                                    let xf_index =
                                        self.get_cell_xf_index(xf_index, row_options, col_num);
                                    let string_index = string_table.shared_string_index(string);
                                    self.write_string_cell(
                                        row_num,
                                        col_num,
                                        &string_index,
                                        &xf_index,
                                    );
                                }
                                CellType::Formula {
                                    formula,
                                    xf_index,
                                    result,
                                } => {
                                    let xf_index =
                                        self.get_cell_xf_index(xf_index, row_options, col_num);
                                    self.write_formula_cell(
                                        row_num, col_num, formula, &xf_index, result,
                                    )
                                }
                                CellType::ArrayFormula {
                                    formula,
                                    xf_index,
                                    result,
                                    is_dynamic,
                                    range,
                                } => {
                                    let xf_index =
                                        self.get_cell_xf_index(xf_index, row_options, col_num);
                                    self.write_array_formula_cell(
                                        row_num, col_num, formula, &xf_index, result, is_dynamic,
                                        range,
                                    )
                                }
                                CellType::Blank { xf_index } => {
                                    let xf_index =
                                        self.get_cell_xf_index(xf_index, row_options, col_num);
                                    self.write_blank_cell(row_num, col_num, &xf_index);
                                }
                                CellType::Boolean { boolean, xf_index } => {
                                    let xf_index =
                                        self.get_cell_xf_index(xf_index, row_options, col_num);
                                    self.write_boolean_cell(row_num, col_num, boolean, &xf_index);
                                }
                            }
                        }
                    }
                    self.writer.xml_end_tag("row");
                } else {
                    self.write_row(row_num, span, row_options, false);
                }
            }
        }
    }

    // Calculate the "spans" attribute of the <row> tag. This is an XLSX
    // optimization and isn't strictly required. However, it makes comparing
    // files easier. The span is the same for each block of 16 rows.
    fn calculate_spans(&mut self) -> HashMap<u32, String> {
        let mut spans: HashMap<RowNum, String> = HashMap::new();
        let mut span_min = COL_MAX;
        let mut span_max = 0;

        for row_num in self.dimensions.row_min..=self.dimensions.row_max {
            if let Some(columns) = self.table.get(&row_num) {
                for col_num in self.dimensions.col_min..=self.dimensions.col_max {
                    match columns.get(&col_num) {
                        Some(_) => {
                            if span_min == COL_MAX {
                                span_min = col_num;
                                span_max = col_num;
                            } else {
                                span_min = cmp::min(span_min, col_num);
                                span_max = cmp::max(span_max, col_num);
                            }
                        }
                        _ => continue,
                    }
                }
            }

            // Store the span range for each block or 16 rows.
            if (row_num + 1) % 16 == 0 || row_num == self.dimensions.row_max {
                let span_index = row_num / 16;
                if span_min != COL_MAX {
                    span_min += 1;
                    span_max += 1;
                    let span_range = format!("{}:{}", span_min, span_max);
                    spans.insert(span_index, span_range);
                    span_min = COL_MAX;
                }
            }
        }

        spans
    }

    // Write the <row> element.
    fn write_row(
        &mut self,
        row_num: RowNum,
        span: Option<&String>,
        row_options: Option<&RowOptions>,
        has_data: bool,
    ) {
        let row_num = format!("{}", row_num + 1);
        let mut attributes = vec![("r", row_num)];

        if let Some(span_range) = span {
            attributes.push(("spans", span_range.clone()));
        }

        if let Some(row_options) = row_options {
            let mut xf_index = row_options.xf_index;

            if xf_index != 0 {
                xf_index = self.global_xf_indices[xf_index as usize];
                attributes.push(("s", xf_index.to_string()));
                attributes.push(("customFormat", "1".to_string()));
            }
            if row_options.height != DEFAULT_ROW_HEIGHT {
                attributes.push(("ht", row_options.height.to_string()));
                attributes.push(("customHeight", "1".to_string()));
            }
        }

        if has_data {
            self.writer.xml_start_tag_attr("row", &attributes);
        } else {
            self.writer.xml_empty_tag_attr("row", &attributes);
        }
    }

    // Write the <c> element for a number.
    fn write_number_cell(&mut self, row: RowNum, col: ColNum, number: &f64, xf_index: &u32) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{}><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            number
        )
        .expect("Couldn't write to file");
    }

    // Write the <c> element for a string.
    fn write_string_cell(&mut self, row: RowNum, col: ColNum, string_index: &u32, xf_index: &u32) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{} t="s"><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            string_index
        )
        .expect("Couldn't write to file");
    }

    // Write the <c> element for a formula.
    fn write_formula_cell(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
        xf_index: &u32,
        result: &str,
    ) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");
        let mut result_type = String::from("");

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        if result.parse::<f64>().is_err() {
            result_type = String::from(r#" t="str""#);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{}{}><f>{}</f><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            result_type,
            crate::xmlwriter::escape_data(formula),
            crate::xmlwriter::escape_data(result),
        )
        .expect("Couldn't write to file");
    }

    // Write the <c> element for an array formula.
    #[allow(clippy::too_many_arguments)]
    fn write_array_formula_cell(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: &str,
        xf_index: &u32,
        result: &str,
        is_dynamic: &bool,
        range: &str,
    ) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");
        let mut cm = String::from("");
        let mut result_type = String::from("");

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        if *is_dynamic {
            cm = String::from(r#" cm="1""#);
        }

        if result.parse::<f64>().is_err() {
            result_type = String::from(r#" t="str""#);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{}{}{}><f t="array" ref="{}">{}</f><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            cm,
            result_type,
            range,
            crate::xmlwriter::escape_data(formula),
            crate::xmlwriter::escape_data(result),
        )
        .expect("Couldn't write to file");
    }

    // Write the <c> element for a blank cell.
    fn write_blank_cell(&mut self, row: RowNum, col: ColNum, xf_index: &u32) {
        let col_name = self.col_to_name(col);

        // Write formatted blank cells and ignore unformatted blank cells (like
        // Excel does).
        if *xf_index > 0 {
            let style = format!(r#" s="{}""#, *xf_index);

            write!(
                &mut self.writer.xmlfile,
                r#"<c r="{}{}"{}/>"#,
                col_name,
                row + 1,
                style
            )
            .expect("Couldn't write to file");
        }
    }

    // Write the <c> element for a boolean cell.
    fn write_boolean_cell(&mut self, row: RowNum, col: ColNum, boolean: &bool, xf_index: &u32) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");
        let boolean = if *boolean { 1 } else { 0 };

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{} t="b"><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            boolean
        )
        .expect("Couldn't write to file");
    }

    // Write the <cols> element.
    fn write_cols(&mut self) {
        if self.changed_cols.is_empty() {
            return;
        }

        self.writer.xml_start_tag("cols");

        // We need to write contiguous equivalent columns as a range with first
        // and last columns, so we convert the HashMap to a sorted vector and
        // iterate over that.
        let changed_cols = self.changed_cols.clone();
        let mut col_options: Vec<_> = changed_cols.iter().collect();
        col_options.sort_by_key(|x| x.0);

        // Remove the first (key, value) tuple in the vector and use it to set
        // the initial/previous properties.
        let first_col_options = col_options.remove(0);
        let mut first_col = first_col_options.0;
        let mut prev_col_options = first_col_options.1;
        let mut last_col = first_col;

        for (col_num, col_options) in col_options.iter() {
            // Check if the column number is contiguous with the previous column
            // and if the format is the same.
            if **col_num == *last_col + 1 && col_options == &prev_col_options {
                last_col = col_num;
            } else {
                // If not write out the current range of columns and start again.
                self.write_col(first_col, last_col, prev_col_options);
                first_col = *col_num;
                last_col = first_col;
                prev_col_options = *col_options;
            }
        }

        // We will exit the previous loop with one unhandled column range.
        self.write_col(first_col, last_col, prev_col_options);

        self.writer.xml_end_tag("cols");
    }

    // Write the <col> element.
    fn write_col(&mut self, first_col: &ColNum, last_col: &ColNum, col_options: &ColOptions) {
        let mut attributes = vec![];
        let first_col = *first_col + 1;
        let last_col = *last_col + 1;
        let mut width = col_options.width;
        let mut xf_index = col_options.xf_index;
        let has_custom_width = width != 8.43;

        // Convert column width from user units to character width.
        if width > 0.0 {
            // Properties for Calibri 11.
            let max_digit_width = 7.0_f64;
            let padding = 5.0_f64;

            if width < 1.0 {
                width = ((width * (max_digit_width + padding)).round() / max_digit_width * 256.0)
                    .floor()
                    / 256.0;
            } else {
                width = (((width * max_digit_width).round() + padding) / max_digit_width * 256.0)
                    .floor()
                    / 256.0;
            }
        }

        attributes.push(("min", first_col.to_string()));
        attributes.push(("max", last_col.to_string()));
        attributes.push(("width", width.to_string()));

        if xf_index > 0 {
            xf_index = self.global_xf_indices[xf_index as usize];
            attributes.push(("style", xf_index.to_string()));
        }

        if has_custom_width {
            attributes.push(("customWidth", "1".to_string()));
        }

        self.writer.xml_empty_tag_attr("col", &attributes);
    }

    // Write the <headerFooter> element.
    fn write_header_footer(&mut self) {
        self.writer.xml_start_tag("headerFooter");

        // Write the oddHeader element.
        if !self.header.is_empty() {
            self.write_odd_header();
        }

        // Write the oddFooter element.
        if !self.footer.is_empty() {
            self.write_odd_footer();
        }

        self.writer.xml_end_tag("headerFooter");
    }

    // Write the <oddHeader> element.
    fn write_odd_header(&mut self) {
        self.writer.xml_data_element("oddHeader", &self.header);
    }

    // Write the <oddFooter> element.
    fn write_odd_footer(&mut self) {
        self.writer.xml_data_element("oddFooter", &self.footer);
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------

// Utility method to strip equal sign and array braces from a formula and
// also expand out future and dynamic array formulas.
fn prepare_formula(formula: &str, expand_future_functions: bool) -> String {
    let mut formula = formula.to_string();

    // Remove array formula braces and the leading = if they exist.
    if formula.starts_with('{') {
        formula.remove(0);
    }
    if formula.starts_with('=') {
        formula.remove(0);
    }
    if formula.ends_with('}') {
        formula.pop();
    }

    // Exit if formula is already expanded by the user.
    if formula.contains("_xlfn.") {
        return formula;
    }

    // Expand dynamic formulas.
    formula = escape_dynamic_formulas1(&formula).into();
    formula = escape_dynamic_formulas2(&formula).into();

    if expand_future_functions {
        formula = escape_future_functions(&formula).into();
    }

    formula
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

// Check of a dynamic function/formula.
fn is_dynamic_function(formula: &str) -> bool {
    lazy_static! {
        static ref DYNAMIC_FUNCTION: Regex = Regex::new(
            r"\b(ANCHORARRAY|FILTER|LAMBDA|LET|RANDARRAY|SEQUENCE|SINGLE|SORTBY|SORT|UNIQUE|XLOOKUP|XMATCH)\("
        )
        .unwrap();
    }
    DYNAMIC_FUNCTION.is_match(formula)
}

#[derive(Clone)]
struct WorksheetDimensions {
    row_min: RowNum,
    col_min: ColNum,
    row_max: RowNum,
    col_max: ColNum,
}

#[derive(Clone)]
struct RowOptions {
    height: f64,
    xf_index: u32,
}

#[derive(Clone, PartialEq)]
struct ColOptions {
    width: f64,
    xf_index: u32,
}

#[derive(Clone)]
enum CellType {
    ArrayFormula {
        formula: String,
        xf_index: u32,
        result: String,
        is_dynamic: bool,
        range: String,
    },
    Blank {
        xf_index: u32,
    },
    Boolean {
        boolean: bool,
        xf_index: u32,
    },
    Formula {
        formula: String,
        xf_index: u32,
        result: String,
    },
    Number {
        number: f64,
        xf_index: u32,
    },
    String {
        string: String,
        xf_index: u32,
    },
}

#[derive(Clone, Copy)]
enum PageView {
    Normal,
    PageLayout,
    PageBreaks,
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::SharedStringsTable;
    use crate::test_functions::xml_to_vec;
    use crate::worksheet::*;
    use crate::XlsxError;
    use assert_float_eq::{afe_is_f64_near, afe_near_error_msg, assert_f64_near};
    use pretty_assertions::assert_eq;
    use std::collections::HashMap;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet::new("".to_string());
        let mut string_table = SharedStringsTable::new();

        worksheet.selected = true;

        worksheet.assemble_xml_file(&mut string_table);

        let got = worksheet.writer.read_to_string();
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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

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
        let mut worksheet = Worksheet::new("".to_string());

        match worksheet.set_name("") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameCannotBeBlank),
        };

        let name = "name_that_is_longer_than_thirty_one_characters".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameLengthExceeded(name)),
        };

        let name = "name_with_special_character_[".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "name_with_special_character_]".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "name_with_special_character_:".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "name_with_special_character_*".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "name_with_special_character_?".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "name_with_special_character_/".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "name_with_special_character_\\".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter(name)),
        };

        let name = "'start with apostrophe".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameStartsOrEndsWithApostrophe(name)),
        };

        let name = "end with apostrophe'".to_string();
        match worksheet.set_name(&name) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameStartsOrEndsWithApostrophe(name)),
        };
    }

    #[test]
    fn check_dimensions() {
        let mut worksheet = Worksheet::new("".to_string());
        let format = Format::default();

        assert_eq!(worksheet.check_dimensions(ROW_MAX, 0), false);
        assert_eq!(worksheet.check_dimensions(0, COL_MAX), false);

        match worksheet.write_string(ROW_MAX, 0, "", &format) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.write_string_only(ROW_MAX, 0, "") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.write_number(ROW_MAX, 0, 0, &format) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.write_number_only(ROW_MAX, 0, 0) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_row_height(ROW_MAX, 20) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_row_height_pixels(ROW_MAX, 20) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_row_format(ROW_MAX, &format) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_column_width(COL_MAX, 20) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_column_width_pixels(COL_MAX, 20) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_column_format(COL_MAX, &format) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };
    }

    #[test]
    fn long_string() {
        let mut worksheet = Worksheet::new("".to_string());
        let chars: [u8; 32_768] = [64; 32_768];
        let long_string = std::str::from_utf8(&chars);

        match worksheet.write_string_only(0, 0, long_string.unwrap()) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::MaxStringLengthExceeded),
        };
    }

    #[test]
    fn date_times() {
        let mut worksheet = Worksheet::new("".to_string());

        // Test date and time
        let datetime = NaiveDate::from_ymd(1899, 12, 31).and_hms_milli(0, 0, 0, 0);
        assert_eq!(0.0, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(1982, 8, 25).and_hms_milli(0, 15, 20, 213);
        assert_eq!(30188.010650613425, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2065, 4, 19).and_hms_milli(0, 16, 48, 290);
        assert_eq!(60376.011670023145, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2147, 12, 15).and_hms_milli(0, 55, 25, 446);
        assert_eq!(90565.038488958337, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2230, 8, 10).and_hms_milli(1, 2, 46, 891);
        assert_eq!(120753.04359827546, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2313, 4, 6).and_hms_milli(1, 4, 15, 597);
        assert_eq!(150942.04462496529, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2395, 11, 30).and_hms_milli(1, 9, 40, 889);
        assert_eq!(181130.04838991899, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2478, 7, 25).and_hms_milli(1, 11, 32, 560);
        assert_eq!(211318.04968240741, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2561, 3, 21).and_hms_milli(1, 30, 19, 169);
        assert_eq!(241507.06272186342, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2643, 11, 15).and_hms_milli(1, 48, 25, 580);
        assert_eq!(271695.07529606484, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2726, 7, 12).and_hms_milli(2, 3, 31, 919);
        assert_eq!(301884.08578609955, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2809, 3, 6).and_hms_milli(2, 11, 11, 986);
        assert_eq!(332072.09111094906, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2891, 10, 31).and_hms_milli(2, 24, 37, 95);
        assert_eq!(362261.10042934027, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(2974, 6, 26).and_hms_milli(2, 35, 7, 220);
        assert_eq!(392449.10772245371, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3057, 2, 19).and_hms_milli(2, 45, 12, 109);
        assert_eq!(422637.1147234838, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3139, 10, 17).and_hms_milli(3, 6, 39, 990);
        assert_eq!(452826.12962951389, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3222, 6, 11).and_hms_milli(3, 8, 8, 251);
        assert_eq!(483014.13065105322, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3305, 2, 5).and_hms_milli(3, 19, 12, 576);
        assert_eq!(513203.13834, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3387, 10, 1).and_hms_milli(3, 29, 42, 574);
        assert_eq!(543391.14563164348, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3470, 5, 27).and_hms_milli(3, 37, 30, 813);
        assert_eq!(573579.15105107636, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3553, 1, 21).and_hms_milli(4, 14, 38, 231);
        assert_eq!(603768.17683137732, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3635, 9, 16).and_hms_milli(4, 16, 28, 559);
        assert_eq!(633956.17810832174, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3718, 5, 13).and_hms_milli(4, 17, 58, 222);
        assert_eq!(664145.17914608796, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3801, 1, 6).and_hms_milli(4, 21, 41, 794);
        assert_eq!(694333.18173372687, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3883, 9, 2).and_hms_milli(4, 56, 35, 792);
        assert_eq!(724522.20596981479, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(3966, 4, 28).and_hms_milli(5, 25, 14, 885);
        assert_eq!(754710.2258667245, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4048, 12, 21).and_hms_milli(5, 26, 5, 724);
        assert_eq!(784898.22645513888, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4131, 8, 18).and_hms_milli(5, 46, 44, 68);
        assert_eq!(815087.24078782403, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4214, 4, 13).and_hms_milli(5, 48, 1, 141);
        assert_eq!(845275.24167987274, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4296, 12, 7).and_hms_milli(5, 53, 52, 315);
        assert_eq!(875464.24574438657, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4379, 8, 3).and_hms_milli(6, 14, 48, 580);
        assert_eq!(905652.26028449077, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4462, 3, 28).and_hms_milli(6, 46, 15, 738);
        assert_eq!(935840.28212659725, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4544, 11, 22).and_hms_milli(7, 31, 20, 407);
        assert_eq!(966029.31343063654, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4627, 7, 19).and_hms_milli(7, 58, 33, 754);
        assert_eq!(996217.33233511576, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4710, 3, 15).and_hms_milli(8, 7, 43, 130);
        assert_eq!(1026406.3386936343, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4792, 11, 7).and_hms_milli(8, 29, 11, 91);
        assert_eq!(1056594.3536005903, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4875, 7, 4).and_hms_milli(9, 8, 15, 328);
        assert_eq!(1086783.3807329629, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(4958, 2, 27).and_hms_milli(9, 30, 41, 781);
        assert_eq!(1116971.3963169097, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5040, 10, 23).and_hms_milli(9, 34, 4, 462);
        assert_eq!(1147159.3986627546, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5123, 6, 20).and_hms_milli(9, 37, 23, 945);
        assert_eq!(1177348.4009715857, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5206, 2, 12).and_hms_milli(9, 37, 56, 655);
        assert_eq!(1207536.4013501736, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5288, 10, 8).and_hms_milli(9, 45, 12, 230);
        assert_eq!(1237725.406391551, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5371, 6, 4).and_hms_milli(9, 54, 14, 782);
        assert_eq!(1267913.412671088, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5454, 1, 28).and_hms_milli(9, 54, 22, 108);
        assert_eq!(1298101.4127558796, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5536, 9, 24).and_hms_milli(10, 1, 36, 151);
        assert_eq!(1328290.4177795255, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5619, 5, 20).and_hms_milli(12, 9, 48, 602);
        assert_eq!(1358478.5068125231, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5702, 1, 14).and_hms_milli(12, 34, 8, 549);
        assert_eq!(1388667.5237100578, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5784, 9, 8).and_hms_milli(12, 56, 6, 495);
        assert_eq!(1418855.5389640625, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5867, 5, 6).and_hms_milli(12, 58, 58, 217);
        assert_eq!(1449044.5409515856, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(5949, 12, 30).and_hms_milli(12, 59, 54, 263);
        assert_eq!(1479232.5416002662, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6032, 8, 24).and_hms_milli(13, 34, 41, 331);
        assert_eq!(1509420.5657561459, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6115, 4, 21).and_hms_milli(13, 58, 28, 601);
        assert_eq!(1539609.5822754744, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6197, 12, 14).and_hms_milli(14, 2, 16, 899);
        assert_eq!(1569797.5849178126, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6280, 8, 10).and_hms_milli(14, 36, 17, 444);
        assert_eq!(1599986.6085352316, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6363, 4, 6).and_hms_milli(14, 37, 57, 451);
        assert_eq!(1630174.60969272, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6445, 11, 30).and_hms_milli(14, 57, 42, 757);
        assert_eq!(1660363.6234115392, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6528, 7, 26).and_hms_milli(15, 10, 48, 307);
        assert_eq!(1690551.6325035533, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6611, 3, 22).and_hms_milli(15, 14, 39, 890);
        assert_eq!(1720739.635183912, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6693, 11, 15).and_hms_milli(15, 19, 47, 988);
        assert_eq!(1750928.6387498612, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6776, 7, 11).and_hms_milli(16, 4, 24, 344);
        assert_eq!(1781116.6697262037, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6859, 3, 7).and_hms_milli(16, 22, 23, 952);
        assert_eq!(1811305.6822216667, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(6941, 10, 31).and_hms_milli(16, 29, 55, 999);
        assert_eq!(1841493.6874536921, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7024, 6, 26).and_hms_milli(16, 58, 20, 259);
        assert_eq!(1871681.7071789235, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7107, 2, 21).and_hms_milli(17, 4, 2, 415);
        assert_eq!(1901870.7111390624, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7189, 10, 16).and_hms_milli(17, 18, 29, 630);
        assert_eq!(1932058.7211762732, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7272, 6, 11).and_hms_milli(17, 47, 21, 323);
        assert_eq!(1962247.7412190163, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7355, 2, 5).and_hms_milli(17, 53, 29, 866);
        assert_eq!(1992435.7454845603, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7437, 10, 2).and_hms_milli(17, 53, 41, 76);
        assert_eq!(2022624.7456143056, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7520, 5, 28).and_hms_milli(17, 55, 6, 44);
        assert_eq!(2052812.7465977315, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7603, 1, 21).and_hms_milli(18, 14, 49, 151);
        assert_eq!(2083000.7602910995, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7685, 9, 16).and_hms_milli(18, 17, 45, 738);
        assert_eq!(2113189.7623349307, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7768, 5, 12).and_hms_milli(18, 29, 59, 700);
        assert_eq!(2143377.7708298611, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7851, 1, 7).and_hms_milli(18, 33, 21, 233);
        assert_eq!(2173566.773162419, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(7933, 9, 2).and_hms_milli(19, 14, 24, 673);
        assert_eq!(2203754.8016744559, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8016, 4, 27).and_hms_milli(19, 17, 12, 816);
        assert_eq!(2233942.8036205554, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8098, 12, 22).and_hms_milli(19, 23, 36, 418);
        assert_eq!(2264131.8080603937, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8181, 8, 17).and_hms_milli(19, 46, 25, 908);
        assert_eq!(2294319.8239109721, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8264, 4, 13).and_hms_milli(20, 7, 47, 314);
        assert_eq!(2324508.8387420601, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8346, 12, 8).and_hms_milli(20, 31, 37, 603);
        assert_eq!(2354696.855296331, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8429, 8, 3).and_hms_milli(20, 39, 57, 770);
        assert_eq!(2384885.8610853008, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8512, 3, 29).and_hms_milli(20, 50, 17, 67);
        assert_eq!(2415073.8682530904, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8594, 11, 22).and_hms_milli(21, 2, 57, 827);
        assert_eq!(2445261.8770581828, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8677, 7, 19).and_hms_milli(21, 23, 5, 519);
        assert_eq!(2475450.8910360998, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8760, 3, 14).and_hms_milli(21, 34, 49, 572);
        assert_eq!(2505638.8991848612, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8842, 11, 8).and_hms_milli(21, 39, 5, 944);
        assert_eq!(2535827.9021521294, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(8925, 7, 4).and_hms_milli(21, 39, 18, 426);
        assert_eq!(2566015.9022965971, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9008, 2, 28).and_hms_milli(21, 46, 7, 769);
        assert_eq!(2596203.9070343636, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9090, 10, 24).and_hms_milli(21, 57, 55, 662);
        assert_eq!(2626392.9152275696, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9173, 6, 19).and_hms_milli(22, 19, 11, 732);
        assert_eq!(2656580.9299968979, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9256, 2, 13).and_hms_milli(22, 23, 51, 376);
        assert_eq!(2686769.9332335186, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9338, 10, 9).and_hms_milli(22, 27, 58, 771);
        assert_eq!(2716957.9360968866, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9421, 6, 5).and_hms_milli(22, 43, 30, 392);
        assert_eq!(2747146.9468795368, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9504, 1, 30).and_hms_milli(22, 48, 25, 834);
        assert_eq!(2777334.9502990046, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9586, 9, 24).and_hms_milli(22, 53, 51, 727);
        assert_eq!(2807522.9540709145, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9669, 5, 20).and_hms_milli(23, 12, 56, 536);
        assert_eq!(2837711.9673210187, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9752, 1, 14).and_hms_milli(23, 15, 54, 109);
        assert_eq!(2867899.9693762613, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9834, 9, 10).and_hms_milli(23, 17, 12, 632);
        assert_eq!(2898088.9702850925, worksheet.datetime_to_excel(datetime));

        let datetime = NaiveDate::from_ymd(9999, 12, 31).and_hms_milli(23, 59, 59, 0);
        assert_eq!(2958465.999988426, worksheet.datetime_to_excel(datetime));

        // Test date only.
        let date = NaiveDate::from_ymd(1899, 12, 31);
        assert_eq!(0.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 1, 1);
        assert_eq!(1.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 2, 27);
        assert_eq!(58.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 2, 28);
        assert_eq!(59.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 3, 1);
        assert_eq!(61.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 3, 2);
        assert_eq!(62.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 3, 11);
        assert_eq!(71.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 4, 8);
        assert_eq!(99.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1900, 9, 12);
        assert_eq!(256.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1901, 5, 3);
        assert_eq!(489.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1901, 10, 13);
        assert_eq!(652.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1902, 2, 15);
        assert_eq!(777.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1902, 6, 6);
        assert_eq!(888.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1902, 9, 25);
        assert_eq!(999.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1902, 9, 27);
        assert_eq!(1001.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1903, 4, 26);
        assert_eq!(1212.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1903, 8, 5);
        assert_eq!(1313.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1903, 12, 31);
        assert_eq!(1461.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1904, 1, 1);
        assert_eq!(1462.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1904, 2, 28);
        assert_eq!(1520.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1904, 2, 29);
        assert_eq!(1521.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1904, 3, 1);
        assert_eq!(1522.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 2, 27);
        assert_eq!(2615.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 2, 28);
        assert_eq!(2616.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 3, 1);
        assert_eq!(2617.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 3, 2);
        assert_eq!(2618.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 3, 3);
        assert_eq!(2619.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 3, 4);
        assert_eq!(2620.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 3, 5);
        assert_eq!(2621.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1907, 3, 6);
        assert_eq!(2622.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 1, 1);
        assert_eq!(36161.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 1, 31);
        assert_eq!(36191.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 2, 1);
        assert_eq!(36192.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 2, 28);
        assert_eq!(36219.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 3, 1);
        assert_eq!(36220.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 3, 31);
        assert_eq!(36250.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 4, 1);
        assert_eq!(36251.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 4, 30);
        assert_eq!(36280.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 5, 1);
        assert_eq!(36281.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 5, 31);
        assert_eq!(36311.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 6, 1);
        assert_eq!(36312.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 6, 30);
        assert_eq!(36341.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 7, 1);
        assert_eq!(36342.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 7, 31);
        assert_eq!(36372.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 8, 1);
        assert_eq!(36373.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 8, 31);
        assert_eq!(36403.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 9, 1);
        assert_eq!(36404.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 9, 30);
        assert_eq!(36433.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 10, 1);
        assert_eq!(36434.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 10, 31);
        assert_eq!(36464.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 11, 1);
        assert_eq!(36465.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 11, 30);
        assert_eq!(36494.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 12, 1);
        assert_eq!(36495.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(1999, 12, 31);
        assert_eq!(36525.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 1, 1);
        assert_eq!(36526.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 1, 31);
        assert_eq!(36556.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 2, 1);
        assert_eq!(36557.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 2, 29);
        assert_eq!(36585.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 3, 1);
        assert_eq!(36586.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 3, 31);
        assert_eq!(36616.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 4, 1);
        assert_eq!(36617.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 4, 30);
        assert_eq!(36646.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 5, 1);
        assert_eq!(36647.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 5, 31);
        assert_eq!(36677.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 6, 1);
        assert_eq!(36678.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 6, 30);
        assert_eq!(36707.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 7, 1);
        assert_eq!(36708.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 7, 31);
        assert_eq!(36738.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 8, 1);
        assert_eq!(36739.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 8, 31);
        assert_eq!(36769.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 9, 1);
        assert_eq!(36770.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 9, 30);
        assert_eq!(36799.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 10, 1);
        assert_eq!(36800.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 10, 31);
        assert_eq!(36830.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 11, 1);
        assert_eq!(36831.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 11, 30);
        assert_eq!(36860.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 12, 1);
        assert_eq!(36861.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2000, 12, 31);
        assert_eq!(36891.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 1, 1);
        assert_eq!(36892.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 1, 31);
        assert_eq!(36922.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 2, 1);
        assert_eq!(36923.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 2, 28);
        assert_eq!(36950.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 3, 1);
        assert_eq!(36951.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 3, 31);
        assert_eq!(36981.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 4, 1);
        assert_eq!(36982.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 4, 30);
        assert_eq!(37011.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 5, 1);
        assert_eq!(37012.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 5, 31);
        assert_eq!(37042.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 6, 1);
        assert_eq!(37043.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 6, 30);
        assert_eq!(37072.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 7, 1);
        assert_eq!(37073.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 7, 31);
        assert_eq!(37103.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 8, 1);
        assert_eq!(37104.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 8, 31);
        assert_eq!(37134.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 9, 1);
        assert_eq!(37135.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 9, 30);
        assert_eq!(37164.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 10, 1);
        assert_eq!(37165.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 10, 31);
        assert_eq!(37195.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 11, 1);
        assert_eq!(37196.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 11, 30);
        assert_eq!(37225.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 12, 1);
        assert_eq!(37226.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2001, 12, 31);
        assert_eq!(37256.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 1, 1);
        assert_eq!(182623.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 1, 31);
        assert_eq!(182653.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 2, 1);
        assert_eq!(182654.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 2, 29);
        assert_eq!(182682.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 3, 1);
        assert_eq!(182683.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 3, 31);
        assert_eq!(182713.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 4, 1);
        assert_eq!(182714.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 4, 30);
        assert_eq!(182743.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 5, 1);
        assert_eq!(182744.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 5, 31);
        assert_eq!(182774.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 6, 1);
        assert_eq!(182775.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 6, 30);
        assert_eq!(182804.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 7, 1);
        assert_eq!(182805.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 7, 31);
        assert_eq!(182835.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 8, 1);
        assert_eq!(182836.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 8, 31);
        assert_eq!(182866.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 9, 1);
        assert_eq!(182867.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 9, 30);
        assert_eq!(182896.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 10, 1);
        assert_eq!(182897.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 10, 31);
        assert_eq!(182927.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 11, 1);
        assert_eq!(182928.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 11, 30);
        assert_eq!(182957.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 12, 1);
        assert_eq!(182958.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(2400, 12, 31);
        assert_eq!(182988.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 1, 1);
        assert_eq!(767011.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 1, 31);
        assert_eq!(767041.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 2, 1);
        assert_eq!(767042.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 2, 29);
        assert_eq!(767070.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 3, 1);
        assert_eq!(767071.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 3, 31);
        assert_eq!(767101.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 4, 1);
        assert_eq!(767102.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 4, 30);
        assert_eq!(767131.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 5, 1);
        assert_eq!(767132.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 5, 31);
        assert_eq!(767162.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 6, 1);
        assert_eq!(767163.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 6, 30);
        assert_eq!(767192.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 7, 1);
        assert_eq!(767193.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 7, 31);
        assert_eq!(767223.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 8, 1);
        assert_eq!(767224.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 8, 31);
        assert_eq!(767254.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 9, 1);
        assert_eq!(767255.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 9, 30);
        assert_eq!(767284.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 10, 1);
        assert_eq!(767285.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 10, 31);
        assert_eq!(767315.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 11, 1);
        assert_eq!(767316.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 11, 30);
        assert_eq!(767345.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 12, 1);
        assert_eq!(767346.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4000, 12, 31);
        assert_eq!(767376.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 1, 1);
        assert_eq!(884254.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 1, 31);
        assert_eq!(884284.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 2, 1);
        assert_eq!(884285.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 2, 28);
        assert_eq!(884312.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 3, 1);
        assert_eq!(884313.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 3, 31);
        assert_eq!(884343.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 4, 1);
        assert_eq!(884344.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 4, 30);
        assert_eq!(884373.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 5, 1);
        assert_eq!(884374.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 5, 31);
        assert_eq!(884404.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 6, 1);
        assert_eq!(884405.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 6, 30);
        assert_eq!(884434.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 7, 1);
        assert_eq!(884435.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 7, 31);
        assert_eq!(884465.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 8, 1);
        assert_eq!(884466.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 8, 31);
        assert_eq!(884496.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 9, 1);
        assert_eq!(884497.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 9, 30);
        assert_eq!(884526.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 10, 1);
        assert_eq!(884527.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 10, 31);
        assert_eq!(884557.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 11, 1);
        assert_eq!(884558.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 11, 30);
        assert_eq!(884587.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 12, 1);
        assert_eq!(884588.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(4321, 12, 31);
        assert_eq!(884618.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 1, 1);
        assert_eq!(2958101.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 1, 31);
        assert_eq!(2958131.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 2, 1);
        assert_eq!(2958132.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 2, 28);
        assert_eq!(2958159.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 3, 1);
        assert_eq!(2958160.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 3, 31);
        assert_eq!(2958190.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 4, 1);
        assert_eq!(2958191.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 4, 30);
        assert_eq!(2958220.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 5, 1);
        assert_eq!(2958221.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 5, 31);
        assert_eq!(2958251.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 6, 1);
        assert_eq!(2958252.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 6, 30);
        assert_eq!(2958281.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 7, 1);
        assert_eq!(2958282.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 7, 31);
        assert_eq!(2958312.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 8, 1);
        assert_eq!(2958313.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 8, 31);
        assert_eq!(2958343.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 9, 1);
        assert_eq!(2958344.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 9, 30);
        assert_eq!(2958373.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 10, 1);
        assert_eq!(2958374.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 10, 31);
        assert_eq!(2958404.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 11, 1);
        assert_eq!(2958405.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 11, 30);
        assert_eq!(2958434.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 12, 1);
        assert_eq!(2958435.0, worksheet.date_to_excel(date));

        let date = NaiveDate::from_ymd(9999, 12, 31);
        assert_eq!(2958465.0, worksheet.date_to_excel(date));

        // Test time only.
        let time = NaiveTime::from_hms_milli(0, 0, 0, 0);
        assert_f64_near!(0.0, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(0, 15, 20, 213);
        assert_f64_near!(1.0650613425925924E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(0, 16, 48, 290);
        assert_f64_near!(1.1670023148148148E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(0, 55, 25, 446);
        assert_f64_near!(3.8488958333333337E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(1, 2, 46, 891);
        assert_f64_near!(4.3598275462962965E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(1, 4, 15, 597);
        assert_f64_near!(4.4624965277777782E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(1, 9, 40, 889);
        assert_f64_near!(4.8389918981481483E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(1, 11, 32, 560);
        assert_f64_near!(4.9682407407407404E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(1, 30, 19, 169);
        assert_f64_near!(6.2721863425925936E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(1, 48, 25, 580);
        assert_f64_near!(7.5296064814814809E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(2, 3, 31, 919);
        assert_f64_near!(8.5786099537037031E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(2, 11, 11, 986);
        assert_f64_near!(9.1110949074074077E-2, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(2, 24, 37, 95);
        assert_f64_near!(0.10042934027777778, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(2, 35, 7, 220);
        assert_f64_near!(0.1077224537037037, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(2, 45, 12, 109);
        assert_f64_near!(0.11472348379629631, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(3, 6, 39, 990);
        assert_f64_near!(0.12962951388888888, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(3, 8, 8, 251);
        assert_f64_near!(0.13065105324074075, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(3, 19, 12, 576);
        assert_f64_near!(0.13833999999999999, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(3, 29, 42, 574);
        assert_f64_near!(0.14563164351851851, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(3, 37, 30, 813);
        assert_f64_near!(0.1510510763888889, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(4, 14, 38, 231);
        assert_f64_near!(0.1768313773148148, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(4, 16, 28, 559);
        assert_f64_near!(0.17810832175925925, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(4, 17, 58, 222);
        assert_f64_near!(0.17914608796296297, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(4, 21, 41, 794);
        assert_f64_near!(0.18173372685185185, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(4, 56, 35, 792);
        assert_f64_near!(0.2059698148148148, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(5, 25, 14, 885);
        assert_f64_near!(0.22586672453703704, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(5, 26, 5, 724);
        assert_f64_near!(0.22645513888888891, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(5, 46, 44, 68);
        assert_f64_near!(0.24078782407407406, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(5, 48, 1, 141);
        assert_f64_near!(0.2416798726851852, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(5, 53, 52, 315);
        assert_f64_near!(0.24574438657407408, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(6, 14, 48, 580);
        assert_f64_near!(0.26028449074074073, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(6, 46, 15, 738);
        assert_f64_near!(0.28212659722222222, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(7, 31, 20, 407);
        assert_f64_near!(0.31343063657407405, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(7, 58, 33, 754);
        assert_f64_near!(0.33233511574074076, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(8, 7, 43, 130);
        assert_f64_near!(0.33869363425925925, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(8, 29, 11, 91);
        assert_f64_near!(0.35360059027777774, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 8, 15, 328);
        assert_f64_near!(0.380732962962963, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 30, 41, 781);
        assert_f64_near!(0.39631690972222228, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 34, 4, 462);
        assert_f64_near!(0.39866275462962958, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 37, 23, 945);
        assert_f64_near!(0.40097158564814817, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 37, 56, 655);
        assert_f64_near!(0.40135017361111114, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 45, 12, 230);
        assert_f64_near!(0.40639155092592594, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 54, 14, 782);
        assert_f64_near!(0.41267108796296298, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(9, 54, 22, 108);
        assert_f64_near!(0.41275587962962962, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(10, 1, 36, 151);
        assert_f64_near!(0.41777952546296299, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(12, 9, 48, 602);
        assert_f64_near!(0.50681252314814818, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(12, 34, 8, 549);
        assert_f64_near!(0.52371005787037039, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(12, 56, 6, 495);
        assert_f64_near!(0.53896406249999995, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(12, 58, 58, 217);
        assert_f64_near!(0.54095158564814816, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(12, 59, 54, 263);
        assert_f64_near!(0.54160026620370372, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(13, 34, 41, 331);
        assert_f64_near!(0.56575614583333333, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(13, 58, 28, 601);
        assert_f64_near!(0.58227547453703699, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(14, 2, 16, 899);
        assert_f64_near!(0.58491781249999997, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(14, 36, 17, 444);
        assert_f64_near!(0.60853523148148148, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(14, 37, 57, 451);
        assert_f64_near!(0.60969271990740748, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(14, 57, 42, 757);
        assert_f64_near!(0.6234115393518519, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(15, 10, 48, 307);
        assert_f64_near!(0.6325035532407407, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(15, 14, 39, 890);
        assert_f64_near!(0.63518391203703706, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(15, 19, 47, 988);
        assert_f64_near!(0.63874986111111109, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(16, 4, 24, 344);
        assert_f64_near!(0.66972620370370362, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(16, 22, 23, 952);
        assert_f64_near!(0.68222166666666662, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(16, 29, 55, 999);
        assert_f64_near!(0.6874536921296297, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(16, 58, 20, 259);
        assert_f64_near!(0.70717892361111112, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(17, 4, 2, 415);
        assert_f64_near!(0.71113906250000003, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(17, 18, 29, 630);
        assert_f64_near!(0.72117627314814825, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(17, 47, 21, 323);
        assert_f64_near!(0.74121901620370367, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(17, 53, 29, 866);
        assert_f64_near!(0.74548456018518516, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(17, 53, 41, 76);
        assert_f64_near!(0.74561430555555563, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(17, 55, 6, 44);
        assert_f64_near!(0.74659773148148145, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(18, 14, 49, 151);
        assert_f64_near!(0.760291099537037, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(18, 17, 45, 738);
        assert_f64_near!(0.76233493055555546, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(18, 29, 59, 700);
        assert_f64_near!(0.77082986111111118, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(18, 33, 21, 233);
        assert_f64_near!(0.77316241898148153, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(19, 14, 24, 673);
        assert_f64_near!(0.80167445601851861, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(19, 17, 12, 816);
        assert_f64_near!(0.80362055555555545, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(19, 23, 36, 418);
        assert_f64_near!(0.80806039351851855, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(19, 46, 25, 908);
        assert_f64_near!(0.82391097222222232, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(20, 7, 47, 314);
        assert_f64_near!(0.83874206018518516, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(20, 31, 37, 603);
        assert_f64_near!(0.85529633101851854, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(20, 39, 57, 770);
        assert_f64_near!(0.86108530092592594, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(20, 50, 17, 67);
        assert_f64_near!(0.86825309027777775, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 2, 57, 827);
        assert_f64_near!(0.87705818287037041, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 23, 5, 519);
        assert_f64_near!(0.891036099537037, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 34, 49, 572);
        assert_f64_near!(0.89918486111111118, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 39, 5, 944);
        assert_f64_near!(0.90215212962962965, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 39, 18, 426);
        assert_f64_near!(0.90229659722222222, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 46, 7, 769);
        assert_f64_near!(0.90703436342592603, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(21, 57, 55, 662);
        assert_f64_near!(0.91522756944444439, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(22, 19, 11, 732);
        assert_f64_near!(0.92999689814814823, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(22, 23, 51, 376);
        assert_f64_near!(0.93323351851851843, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(22, 27, 58, 771);
        assert_f64_near!(0.93609688657407408, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(22, 43, 30, 392);
        assert_f64_near!(0.94687953703703709, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(22, 48, 25, 834);
        assert_f64_near!(0.95029900462962968, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(22, 53, 51, 727);
        assert_f64_near!(0.95407091435185187, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(23, 12, 56, 536);
        assert_f64_near!(0.96732101851851848, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(23, 15, 54, 109);
        assert_f64_near!(0.96937626157407408, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(23, 17, 12, 632);
        assert_f64_near!(0.97028509259259266, worksheet.time_to_excel(time));

        let time = NaiveTime::from_hms_milli(23, 59, 59, 999);
        assert_f64_near!(0.99999998842592586, worksheet.time_to_excel(time));
    }
}
