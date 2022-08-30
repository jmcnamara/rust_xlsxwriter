// worksheet - A module for creating the Excel sheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
use std::cmp;
use std::collections::HashMap;
use std::io::Write;
use std::mem;

use regex::Regex;

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
/// such as writing data to cells or formatting worksheet layout.
///
/// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/demo.png">
///
/// # Examples
///
/// Sample code to generate the Excel file shown above.
///
/// ```rust
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("demo.xlsx");
///
///     // Create some formats to use in the worksheet.
///     let bold_format = Format::new().set_bold();
///     let decimal_format = Format::new().set_num_format("0.000");
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
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
///     workbook.close()?;
///
///     Ok(())
/// }
///
/// ```
pub struct Worksheet {
    pub(crate) writer: XMLWriter,
    pub(crate) name: String,
    pub(crate) selected: bool,
    pub(crate) uses_string_table: bool,
    table: HashMap<RowNum, HashMap<ColNum, CellType>>,
    col_names: HashMap<ColNum, String>,
    dimensions: WorksheetDimensions,
    pub(crate) xf_formats: Vec<Format>,
    xf_indices: HashMap<String, u32>,
    global_xf_indices: Vec<u32>,
    changed_rows: HashMap<RowNum, RowOptions>,
    changed_cols: HashMap<ColNum, ColOptions>,
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
            table,
            col_names,
            dimensions,
            xf_formats: vec![default_format],
            xf_indices,
            global_xf_indices: vec![],
            changed_rows,
            changed_cols,
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_name.png">
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
        if name.len() > 31 {
            return Err(XlsxError::SheetnameLengthExceeded);
        }

        // Check that sheetname doesn't contain any invalid characters.
        let re = Regex::new(r"[\[\]:*?/\\]").unwrap();
        if re.is_match(name) {
            return Err(XlsxError::SheetnameContainsInvalidCharacter);
        }

        // Check that sheetname doesn't start or end with an apostrophe.
        if name.starts_with('\'') || name.ends_with('\'') {
            return Err(XlsxError::SheetnameStartsOrEndsWithApostrophe);
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_write_number.png">
    ///
    ///
    pub fn write_number<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: T,
        format: &Format,
    ) -> Result<(), XlsxError>
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_write_number_only.png">
    ///
    pub fn write_number_only<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: T,
    ) -> Result<(), XlsxError>
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_write_string.png">
    ///
    pub fn write_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        format: &Format,
    ) -> Result<(), XlsxError> {
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
    /// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_write_string_only.png">
    ///
    pub fn write_string_only(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
    ) -> Result<(), XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string, None)
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_row_height.png">
    ///
    pub fn set_row_height<T>(&mut self, row: RowNum, height: T) -> Result<(), XlsxError>
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

        Ok(())
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_row_height.png">
    ///
    pub fn set_row_height_pixels(&mut self, row: RowNum, height: u16) -> Result<(), XlsxError> {
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_row_format.png">
    ///
    pub fn set_row_format(&mut self, row: RowNum, format: &Format) -> Result<(), XlsxError> {
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

        Ok(())
    }

    /// Set the width for worksheet columns.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// column or a range of columns.
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
    /// * `first_col` - The zero indexed column number of the first column in
    ///   the range.
    /// * `last_col` - The zero indexed column number of the last column in the
    ///   range. Set this value to the same as `first_col` for a single column.
    /// * `width` - The row width in character units.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - First or last column exceeds
    ///   Excel's worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the width of columns in Excel.
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
    ///     worksheet.write_string_only(0, 2, "Wider")?;
    ///     worksheet.write_string_only(0, 4, "Narrower")?;
    ///
    ///     // Set the column width in Excel character units.
    ///     worksheet.set_column_width(2, 2, 16)?; // Single column.
    ///     worksheet.set_column_width(4, 5, 4)?;  // Column range.
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
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_column_width.png">
    ///
    pub fn set_column_width<T>(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        width: T,
    ) -> Result<(), XlsxError>
    where
        T: Into<f64>,
    {
        let mut first_col = first_col;
        let mut last_col = last_col;
        let width = width.into();

        // Ensure columns are in the right order.
        if first_col > last_col {
            (first_col, last_col) = (last_col, first_col);
        }

        // Check columns are in the allowed range without updating dimensions.
        if first_col >= COL_MAX || last_col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing col metadata object or create a new one.
        for col_num in first_col..=last_col {
            match self.changed_cols.get_mut(&col_num) {
                Some(col_options) => col_options.width = width,
                None => {
                    let col_options = ColOptions { width, xf_index: 0 };
                    self.changed_cols.insert(col_num, col_options);
                }
            }
        }
        Ok(())
    }

    /// Set the width for worksheet columns in pixels.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// column or a range of columns.
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
    /// * `first_col` - The zero indexed column number of the first column in
    ///   the range.
    /// * `last_col` - The zero indexed column number of the last column in the
    ///   range. Set this value to the same as `first_col` for a single column.
    /// * `width` - The row width in pixels.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - First or last column exceeds
    ///   Excel's worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the width of columns in Excel in
    /// pixels.
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
    ///     worksheet.write_string_only(0, 2, "Wider")?;
    ///     worksheet.write_string_only(0, 4, "Narrower")?;
    ///
    ///     // Set the column width in pixels.
    ///     worksheet.set_column_width_pixels(2, 2, 117)?; // Single column.
    ///     worksheet.set_column_width_pixels(4, 5, 33)?;  // Column range.
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
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_column_width.png">
    ///
    pub fn set_column_width_pixels(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        width: u16,
    ) -> Result<(), XlsxError> {
        // Properties for Calibri 11.
        let max_digit_width = 7.0_f64;
        let padding = 5.0_f64;
        let mut width = width as f64;

        if width < 12.0 {
            width /= max_digit_width + padding;
        } else {
            width = (width - padding) / max_digit_width
        }

        self.set_column_width(first_col, last_col, width)
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
    /// * `first_col` - The zero indexed column number of the first column in
    ///   the range.
    /// * `last_col` - The zero indexed column number of the last column in the
    ///   range. Set this value to the same as `first_col` for a single column.
    /// * `format` - The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format for a column in Excel.
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
    ///     // Set the column format.
    ///     worksheet.set_column_format(1, 1, &red_format)?;
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
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/worksheet_set_column_format.png">
    ///
    pub fn set_column_format(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        format: &Format,
    ) -> Result<(), XlsxError> {
        let mut first_col = first_col;
        let mut last_col = last_col;

        // Ensure columns are in the right order.
        if first_col > last_col {
            (first_col, last_col) = (last_col, first_col);
        }

        // Set a suitable row range for the dimension check/set.
        let min_row = if self.dimensions.row_min != ROW_MAX {
            self.dimensions.row_min
        } else {
            0
        };

        // Check columns are in the allowed range.
        if !self.check_dimensions(min_row, first_col) {
            return Err(XlsxError::RowColumnLimitError);
        }
        if !self.check_dimensions(min_row, last_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_index(format);

        // Update an existing col metadata object or create a new one.
        for col_num in first_col..=last_col {
            match self.changed_cols.get_mut(&col_num) {
                Some(col_options) => col_options.xf_index = xf_index,
                None => {
                    let col_options = ColOptions {
                        width: DEFAULT_COL_WIDTH,
                        xf_index,
                    };
                    self.changed_cols.insert(col_num, col_options);
                }
            }
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    // Store a number cell.
    fn store_number(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: f64,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
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

        Ok(())
    }

    // Writer a unformatted string to a cell.
    fn store_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        //  Check that the string is < Excel limit of 32767 chars.
        if string.len() as u16 > MAX_STRING_LEN {
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

        Ok(())
    }

    // Insert a cell value into the worksheet table data structure.
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
    // indexes will be replaced by global/worksheet indices before the worksheet
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

    // Set the mapping between the local format indices and the global/worksheet
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

        if self.selected {
            attributes.push(("tabSelected", "1".to_string()));
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
            ("left", "0.7".to_string()),
            ("right", "0.7".to_string()),
            ("top", "0.75".to_string()),
            ("bottom", "0.75".to_string()),
            ("header", "0.3".to_string()),
            ("footer", "0.3".to_string()),
        ];

        self.writer.xml_empty_tag_attr("pageMargins", &attributes);
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

    // Write the <cols> element.
    fn write_cols(&mut self) {
        if self.changed_cols.is_empty() {
            return;
        }

        self.writer.xml_start_tag("cols");

        // We need to write contiguous equivalent columns as a range with first
        // and last columns, so we convert the hashmap to a sorted vector and
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
}

// -----------------------------------------------------------------------
// Helper enums/structs
// -----------------------------------------------------------------------

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
    Number { number: f64, xf_index: u32 },
    String { string: String, xf_index: u32 },
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

        match worksheet.set_name("name_that_is_longer_than_thirty_one_characters") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameLengthExceeded),
        };

        match worksheet.set_name("name_with_special_character_[") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("name_with_special_character_]") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("name_with_special_character_:") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("name_with_special_character_*") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("name_with_special_character_?") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("name_with_special_character_/") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("name_with_special_character_\\") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameContainsInvalidCharacter),
        };

        match worksheet.set_name("'start with apostrophe") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameStartsOrEndsWithApostrophe),
        };

        match worksheet.set_name("end with apostrophe'") {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::SheetnameStartsOrEndsWithApostrophe),
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

        match worksheet.set_column_width(COL_MAX, 0, 20) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_column_width_pixels(0, COL_MAX, 20) {
            Ok(_) => assert!(false),
            Err(err) => assert_eq!(err, XlsxError::RowColumnLimitError),
        };

        match worksheet.set_column_format(COL_MAX, COL_MAX, &format) {
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
}
