// format - A module for representing Excel cell formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

#[derive(Clone)]
/// The Format struct is used to defined cell working for data in a worksheet.
///
/// The properties of a cell that can be formatted include: fonts, colors,
/// patterns, borders, alignment and number formatting.
///
/// TODO: add image
///
/// TODO: add code
///
/// # Creating and using a Format object
///
/// Formats are created by calling the `Format::new()` method and properties as
/// set using the various methods shown is this section of the document. Once
/// the Format has been created it can be passed to one of the worksheet
/// `write_*()` methods. Multiple properties can be set by chaining them
/// together:
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file.
/// #     let mut workbook = Workbook::new("formats.xlsx");
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
///      // Create a new format and set some properties.
///     let format = Format::new()
///         .set_bold()
///         .set_italic()
///         .set_font_color(XlsxColor::Red);
///
///     worksheet.write_string(0, 0, "Hello", &format)?;
///
/// #     workbook.close()?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_create.png">
///
/// Formats can be cloned in the usual way:
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file.
/// #     let mut workbook = Workbook::new("formats.xlsx");
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Create a new format and set some properties.
///     let format1 = Format::new()
///         .set_bold();
///
///     // Clone a new format and set some properties.
///     let format2 = format1.clone()
///         .set_font_color(XlsxColor::Blue);
///
///     worksheet.write_string(0, 0, "Hello", &format1)?;
///     worksheet.write_string(1, 0, "Hello", &format2)?;
///
/// #     workbook.close()?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_clone.png">
///
/// TODO: Merging
///
/// # Format methods and Format properties
///
/// The following table shows the Excel format categories and the equivalent
/// `rust_xlsxwriter` Format method:
///
/// | Category        | Description          |  Method Name                               |
/// | :-------------- | :------------------- |  :---------------------------------------- |
/// | **Font**        | Font type            |  format_set_font_name()                    |
/// |                 | Font size            |  format_set_font_size()                    |
/// |                 | Font color           |  format_set_font_color()                   |
/// |                 | Bold                 |  [`set_bold()`](Format::set_bold())        |
/// |                 | Italic               |  format_set_italic()                       |
/// |                 | Underline            |  format_set_underline()                    |
/// |                 | Strikeout            |  format_set_font_strikeout()               |
/// |                 | Super/Subscript      |  format_set_font_script()                  |
/// | **Number**      | Numeric format       |  format_set_num_format()                   |
/// | **Protection**  | Unlock cells         |  format_set_unlocked()                     |
/// |                 | Hide formulas        |  format_set_hidden()                       |
/// | **Alignment**   | Horizontal align     |  format_set_align()                        |
/// |                 | Vertical align       |  format_set_align()                        |
/// |                 | Rotation             |  format_set_rotation()                     |
/// |                 | Text wrap            |  format_set_text_wrap()                    |
/// |                 | Indentation          |  format_set_indent()                       |
/// |                 | Shrink to fit        |  format_set_shrink()                       |
/// | **Pattern**     | Cell pattern         |  format_set_pattern()                      |
/// |                 | Background color     |  format_set_bg_color()                     |
/// |                 | Foreground color     |  format_set_fg_color()                     |
/// | **Border**      | Cell border          |  format_set_border()                       |
/// |                 | Bottom border        |  format_set_bottom()                       |
/// |                 | Top border           |  format_set_top()                          |
/// |                 | Left border          |  format_set_left()                         |
/// |                 | Right border         |  format_set_right()                        |
/// |                 | Border color         |  format_set_border_color()                 |
/// |                 | Bottom color         |  format_set_bottom_color()                 |
/// |                 | Top color            |  format_set_top_color()                    |
/// |                 | Left color           |  format_set_left_color()                   |
/// |                 | Right color          |  format_set_right_color()                  |
///
/// # Format Colors
///
/// Format property colors are specified the [`XlsxColor`] enum using a Html
/// style RGB integer value or a limited number of defined colors:
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("colors.xlsx");
///
///     let format1 = Format::new().set_font_color(XlsxColor::Red);
///     let format2 = Format::new().set_font_color(XlsxColor::Green);
///     let format3 = Format::new().set_font_color(XlsxColor::RGB(0x4F026A));
///     let format4 = Format::new().set_font_color(XlsxColor::RGB(0x73CC5F));
///     let format5 = Format::new().set_font_color(XlsxColor::RGB(0xFFACFF));
///     let format6 = Format::new().set_font_color(XlsxColor::RGB(0xCC7E16));
///
///     let worksheet = workbook.add_worksheet();
///     worksheet.write_string(0, 0, "Red", &format1)?;
///     worksheet.write_string(1, 0, "Green", &format2)?;
///     worksheet.write_string(2, 0, "#4F026A", &format3)?;
///     worksheet.write_string(3, 0, "#73CC5F", &format4)?;
///     worksheet.write_string(4, 0, "#FFACFF", &format5)?;
///     worksheet.write_string(5, 0, "#CC7E16", &format6)?;
///
/// #     workbook.close()?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/enum_xlsxcolor.png">
///
/// # Format Defaults
///
/// The default Excel 365 cell format is a font setting of Calibri size 11 with
/// all other properties turned off.
///
/// It is occasionally useful to use a default format with a method that
/// requires a format but where you don't actually want to change the
/// formatting.
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file.
/// #     let mut workbook = Workbook::new("formats.xlsx");
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
///      // Create a new default format.
///      let format = Format::default();
///
///      // These methods calls are equivalent.
///      worksheet.write_string_only(0, 0, "Hello")?;
///      worksheet.write_string     (1, 0, "Hello", &format)?;
/// #
/// #     workbook.close()?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_default.png">
///
///
/// # Number Format Categories
///
/// The [`set_num_format()`](Format::set_num_format) method is used to set the
/// number format for numbers used with
/// [`write_number()`](super::Worksheet::write_number()):
///
/// ```
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("currency_format.xlsx");
///
///     // Add a worksheet.
///     let worksheet = workbook.add_worksheet();
///
///     // Add a format.
///     let currency_format = Format::new().set_num_format("$#,##0.00");
///
///     worksheet.write_number(0, 0, 1234.56, &currency_format)?;
///
///     workbook.close()?;
///
///     Ok(())
/// }
/// ```
///
/// If the number format you use is the same as one of Excel's built in number
/// formats then it will have a number category other than "General" or
/// "Number". The Excel number categories are:
///
/// - General
/// - Number
/// - Currency
/// - Accounting
/// - Date
/// - Time",
/// - Percentage
/// - Fraction
/// - Scientific
/// - Text
/// - Custom.
///
/// In the case of the example above the formatted output shows up as a Number
/// category:
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency1.png">
///
/// If we wanted to have the number format display as a different category, such
/// as Currency, then would need to match the number format string used in the
/// code with the number format used by Excel. The easiest way to do this is to
/// open the Number Formatting dialog in Excel and set the required format:
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency2.png">
///
/// Then, while still in the dialog, change to Custom. The format displayed is
/// the format used by Excel.
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency3.png">
///
/// If we put the format that we found (`"[$$-409]#,##0.00"`) into our previous
/// example and rerun it we will get a number format in the Currency category:
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #    // Create a new Excel file.
/// #    let mut workbook = Workbook::new("currency_format.xlsx");
/// #
/// #    // Add a worksheet.
/// #    let worksheet = workbook.add_worksheet();
/// #
/// #    // Add a format.
///     let currency_format = Format::new().set_num_format("[$$-409]#,##0.00");
///
///     worksheet.write_number(0, 0, 1234.56, &currency_format)?;
///
/// #   workbook.close()?;
/// #
/// #   Ok(())
/// # }
/// ```
///
/// That give us the following updated output. Note that the number category is
/// now shown as Currency:
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency4.png">
///
/// The same process can be used to find format strings for "Date" or
/// "Accountancy" formats.
///
/// # Number Formats in different locales
///
/// As shown in the previous section the `format_set_num_format()` method is
/// used to set the number format for `rust_xlsxwriter` formats. A common use
/// case is to set a number format with a "grouping/thousands" separator and a
/// "decimal" point:
///
/// ```
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("number_format.xlsx");
///
///
///     // Add a worksheet.
///     let worksheet = workbook.add_worksheet();
///
///     // Add a format.
///     let currency_format = Format::new().set_num_format("#,##0.00");
///
///     worksheet.write_number(0, 0, 1234.56, &currency_format)?;
///
///     workbook.close()?;
///
///     Ok(())
/// }
/// ```
///
/// In the US locale (and some others) where the number "grouping/thousands"
/// separator is `","` and the "decimal" point is `"."` which would be shown in
/// Excel as:
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency5.png">
///
/// In other locales these values may be reversed or different. They are
/// generally set in the "Region" settings of Windows or Mac OS. Excel handles
/// this by storing the number format in the file format in the US locale, in
/// this case `#,##0.00`, but renders it according to the regional settings of
/// the host OS. For example, here is the same, unmodified, output file shown
/// above in a German locale:
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency6.png">
///
/// And here is the same file in a Russian locale. Note the use of a space as
/// the "grouping/thousands" separator:
///
/// <img
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_currency7.png">
///
/// In order to replicate Excel's behavior all XlsxWriter programs should use US
/// locale formatting which will then be rendered in the settings of your host
/// OS.
///

pub struct Format {
    pub(crate) xf_index: u32,
    pub(crate) font_index: u16,
    pub(crate) has_font: bool,

    // Number properties.
    pub(crate) num_format: String,
    pub(crate) num_format_index: u16,

    // Font properties.
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: XlsxUnderline,
    pub(crate) font_name: String,
    pub(crate) font_size: f64,
    pub(crate) font_color: XlsxColor,
    pub(crate) font_strikeout: bool,
    pub(crate) font_script: XlsxScript,
    pub(crate) font_family: u8,
    pub(crate) font_charset: u8,
    pub(crate) font_scheme: String,
    pub(crate) font_condense: bool,
    pub(crate) font_extend: bool,
    pub(crate) theme: u8,

    // Alignment properties.
    pub(crate) text_horizontal_align: u8,
    pub(crate) text_wrap: bool,
    pub(crate) text_vertical_align: u8,
    pub(crate) text_justify_last: bool,
    pub(crate) rotation: u16,
    pub(crate) indent: u8,
    pub(crate) shrink: bool,
    pub(crate) reading_order: u8,

    // Border properties.
    pub(crate) bottom: u8,
    pub(crate) top: u8,
    pub(crate) left: u8,
    pub(crate) right: u8,
    pub(crate) diagonal_border: u8,
    pub(crate) diagonal_type: u8,
    pub(crate) bottom_color: XlsxColor,
    pub(crate) top_color: XlsxColor,
    pub(crate) left_color: XlsxColor,
    pub(crate) right_color: XlsxColor,
    pub(crate) diagonal_color: XlsxColor,

    // Fill properties.
    pub(crate) foreground_color: XlsxColor,
    pub(crate) background_color: XlsxColor,
    pub(crate) pattern: u8,

    // Protection properties.
    pub(crate) hidden: bool,
    pub(crate) locked: bool,
}

impl Default for Format {
    fn default() -> Self {
        Self::new()
    }
}

impl Format {
    /// Create a new Format object.
    ///
    /// Create a new Format object to use with worksheet formatting.
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a new format.
    ///
    ///
    /// ```
    /// use rust_xlsxwriter::Format;
    ///
    /// # #[allow(unused_variables)]
    /// fn main() {
    ///
    ///     let format = Format::new();
    ///
    /// }
    /// ```
    pub fn new() -> Format {
        Format {
            xf_index: 0,
            font_index: 0,
            has_font: false,

            num_format: "".to_string(),
            num_format_index: 0,
            bold: false,
            italic: false,
            underline: XlsxUnderline::None,
            font_name: "Calibri".to_string(),
            font_size: 11.0,
            font_color: XlsxColor::Automatic,
            font_strikeout: false,
            font_script: XlsxScript::None,
            font_family: 2,
            font_charset: 0,
            font_scheme: "".to_string(),
            font_condense: false,
            font_extend: false,
            theme: 0,
            hidden: false,
            locked: true,
            text_horizontal_align: 0,
            text_wrap: false,
            text_vertical_align: 0,
            text_justify_last: false,
            rotation: 0,
            foreground_color: XlsxColor::Automatic,
            background_color: XlsxColor::Automatic,
            pattern: 0,
            bottom: 0,
            top: 0,
            left: 0,
            right: 0,
            diagonal_border: 0,
            diagonal_type: 0,
            bottom_color: XlsxColor::Automatic,
            top_color: XlsxColor::Automatic,
            left_color: XlsxColor::Automatic,
            right_color: XlsxColor::Automatic,
            diagonal_color: XlsxColor::Automatic,
            indent: 0,
            shrink: false,
            reading_order: 0,
        }
    }

    // -----------------------------------------------------------------------
    // Crate private methods.
    // -----------------------------------------------------------------------

    pub(crate) fn set_xf_index(&mut self, index: u32) {
        self.xf_index = index;
    }

    pub(crate) fn set_font_index(&mut self, font_index: u16, has_font: bool) {
        self.font_index = font_index;
        self.has_font = has_font;
    }

    pub(crate) fn format_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}",
            self.alignment_key(),
            self.border_key(),
            self.fill_key(),
            self.font_key(),
            self.hidden,
            self.locked,
            self.num_format,
            self.num_format_index,
        )
    }

    pub(crate) fn font_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}",
            self.bold,
            self.font_charset,
            self.font_color.hex_value(),
            self.font_condense,
            self.font_extend,
            self.font_family,
            self.font_name,
            self.font_scheme,
            self.font_script as u8,
            self.font_size,
            self.font_strikeout,
            self.italic,
            self.theme,
            self.underline as u8,
        )
    }

    pub(crate) fn border_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}",
            self.bottom,
            self.bottom_color.hex_value(),
            self.diagonal_border,
            self.diagonal_color.hex_value(),
            self.diagonal_type,
            self.left,
            self.left_color.hex_value(),
            self.right,
            self.right_color.hex_value(),
            self.top,
            self.top_color.hex_value(),
        )
    }

    pub(crate) fn fill_key(&self) -> String {
        format!(
            "{}:{}:{}",
            self.background_color.hex_value(),
            self.foreground_color.hex_value(),
            self.pattern,
        )
    }

    pub(crate) fn alignment_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}",
            self.indent,
            self.reading_order,
            self.rotation,
            self.shrink,
            self.text_horizontal_align,
            self.text_vertical_align,
            self.text_justify_last,
            self.text_wrap,
        )
    }

    pub(crate) fn set_num_format_index_u16(&mut self, num_format_index: u16) {
        self.num_format_index = num_format_index;
    }

    // -----------------------------------------------------------------------
    // Public methods.
    // -----------------------------------------------------------------------

    /// Set the number format for a Format.
    ///
    /// This method is used to define the numerical format of a number in Excel.
    /// It controls whether a number is displayed as an integer, a floating
    /// point number, a date, a currency value or some other user defined
    /// format.
    ///
    /// See also [Number Format Categories] and [Number Formats in different locales].
    ///
    /// [Number Format Categories]: struct.Format.html#number-format-categories
    /// [Number Formats in different locales]: struct.Format.html#number-formats-in-different-locales
    ///
    /// # Arguments
    ///
    /// * `num_format` - The number format property.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting different types of Excel
    /// number formatting.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_num_format("0.00");
    ///     let format2 = Format::new().set_num_format("0.000");
    ///     let format3 = Format::new().set_num_format("#,##0");
    ///     let format4 = Format::new().set_num_format("#,##0.00");
    ///     let format5 = Format::new().set_num_format("mm/dd/yy");
    ///     let format6 = Format::new().set_num_format("mmm d yyyy");
    ///     let format7 = Format::new().set_num_format("d mmmm yyyy");
    ///     let format8 = Format::new().set_num_format("dd/mm/yyyy hh:mm AM/PM");
    ///
    ///     worksheet.write_number(0, 0, 3.1415926, &format1)?;
    ///     worksheet.write_number(1, 0, 3.1415926, &format2)?;
    ///     worksheet.write_number(2, 0, 1234.56,   &format3)?;
    ///     worksheet.write_number(3, 0, 1234.56,   &format4)?;
    ///     worksheet.write_number(4, 0, 44927.521, &format5)?;
    ///     worksheet.write_number(5, 0, 44927.521, &format6)?;
    ///     worksheet.write_number(6, 0, 44927.521, &format7)?;
    ///     worksheet.write_number(7, 0, 44927.521, &format8)?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Note how the numbers above have been displayed by Excel in the output
    /// file according to the given number format:
    ///
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_num_format.png">
    ///
    pub fn set_num_format(mut self, num_format: &str) -> Format {
        self.num_format = num_format.to_string();
        self
    }

    /// Set the number format for a Format using a legacy format index.
    ///
    /// This method is similar to [`set_num_format()`](Format::set_num_format)
    /// except that it uses an index to a limited number of Excel's built-in,
    /// and legacy, number formats.
    ///
    /// Unless you need to specifically access one of Excel's built-in number
    /// formats the [`set_num_format()`](Format::set_num_format) method is a
    /// better solution. This method is mainly included for backward
    /// compatibility and completeness.
    ///
    /// The Excel built-in number formats as shown in the table below:
    ///
    /// | Index | Format String                                        |
    /// | :---- | :--------------------------------------------------- |
    /// | 1     | `0`                                                  |
    /// | 2     | `0.00`                                               |
    /// | 3     | `#,##0`                                              |
    /// | 4     | `#,##0.00`                                           |
    /// | 5     | `($#,##0_);($#,##0)`                                 |
    /// | 6     | `($#,##0_);[Red]($#,##0)`                            |
    /// | 7     | `($#,##0.00_);($#,##0.00)`                           |
    /// | 8     | `($#,##0.00_);[Red]($#,##0.00)`                      |
    /// | 9     | `0%`                                                 |
    /// | 10    | `0.00%`                                              |
    /// | 11    | `0.00E+00`                                           |
    /// | 12    | `# ?/?`                                              |
    /// | 13    | `# ??/??`                                            |
    /// | 14    | `m/d/yy`                                             |
    /// | 15    | `d-mmm-yy`                                           |
    /// | 16    | `d-mmm`                                              |
    /// | 17    | `mmm-yy`                                             |
    /// | 18    | `h:mm AM/PM`                                         |
    /// | 19    | `h:mm:ss AM/PM`                                      |
    /// | 20    | `h:mm`                                               |
    /// | 21    | `h:mm:ss`                                            |
    /// | 22    | `m/d/yy h:mm`                                        |
    /// | ...   | ...                                                  |
    /// | 37    | `(#,##0_);(#,##0)`                                   |
    /// | 38    | `(#,##0_);[Red](#,##0)`                              |
    /// | 39    | `(#,##0.00_);(#,##0.00)`                             |
    /// | 40    | `(#,##0.00_);[Red](#,##0.00)`                        |
    /// | 41    | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`            |
    /// | 42    | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`         |
    /// | 43    | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`    |
    /// | 44    | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` |
    /// | 45    | `mm:ss`                                              |
    /// | 46    | `[h]:mm:ss`                                          |
    /// | 47    | `mm:ss.0`                                            |
    /// | 48    | `##0.0E+0`                                           |
    /// | 49    | `@`                                                  |
    ///
    /// Notes:
    ///
    ///  - Numeric formats 23 to 36 are not documented by Microsoft and may
    ///    differ in international versions. The listed date and currency
    ///    formats may also vary depending on system settings.
    ///  - The dollar sign in the above format appears as the defined local
    ///    currency symbol.
    ///  - These formats can also be set via format_set_num_format().
    ///  - See also formats_categories.
    ///
    /// # Arguments
    ///
    /// * `num_format_index` - The index to one of the inbuilt formats shown in
    ///   the table above.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the bold property for a
    /// format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_num_format_index(15);
    ///
    ///     worksheet.write_number(0, 0, 44927.521, &format)?;
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
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_num_format_index.png">
    ///
    pub fn set_num_format_index(mut self, num_format_index: u8) -> Format {
        self.num_format_index = num_format_index as u16;
        self
    }

    /// Set the bold property for a Format font.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the bold property for a format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_bold();
    ///
    ///     worksheet.write_string(0, 0, "Hello", &format)?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_bold.png">
    ///
    pub fn set_bold(mut self) -> Format {
        self.bold = true;
        self
    }

    /// Set the italic property for the Format font.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the italic property for a format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_italic();
    ///
    ///     worksheet.write_string(0, 0, "Hello", &format)?;
    /// #
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_italic.png">
    ///
    pub fn set_italic(mut self) -> Format {
        self.italic = true;
        self
    }

    /// Set the color property for the Format font.
    ///
    /// The `set_font_color()` method is used to set the font color in a
    /// cell. To set the color of a cell background use the `set_bg_color()` and
    /// `set_pattern()` methods.
    ///
    /// # Arguments
    ///
    /// * `font_color` - The font color property defined by a [`XlsxColor`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the italic property for a
    /// format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_color(XlsxColor::Red);
    ///
    ///     worksheet.write_string(0, 0, "Wheelbarrow", &format)?;
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
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_font_color.png">
    ///
    pub fn set_font_color(mut self, font_color: XlsxColor) -> Format {
        if !font_color.is_valid() {
            return self;
        }

        self.font_color = font_color;
        self
    }

    /// Set the Format font name property.
    ///
    /// Set the font for a cell format. Excel can only display fonts that are
    /// installed on the system that it is running on. Therefore it is generally
    /// best to use standard Excel fonts.
    ///
    /// # Arguments
    ///
    /// * `font_name` - The font name property.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the font name/type for a
    /// format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_name("Avenir Black Oblique");
    ///
    ///     worksheet.write_string(0, 0, "Avenir Black Oblique", &format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_font_name.png">
    ///
    pub fn set_font_name(mut self, font_name: &str) -> Format {
        self.font_name = font_name.to_string();
        self
    }

    /// Set the Format font size property.
    ///
    /// Set the font size of the cell format. The size is generally an integer
    /// value but Excel allows x.5 values (hence the property is a f64 or
    /// types that can convert [`Into`] a f64).
    ///
    /// Excel adjusts the height of a row to accommodate the largest font size
    /// in the row.
    ///
    /// # Arguments
    ///
    /// * `font_size` - The font size property.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the font size for a format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_size(30);
    ///
    ///     worksheet.write_string(0, 0, "Font Size 30", &format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_font_size.png">
    ///
    pub fn set_font_size<T>(mut self, font_size: T) -> Format
    where
        T: Into<f64>,
    {
        self.font_size = font_size.into();
        self
    }

    /// Set the Format font scheme property.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    pub fn set_font_scheme(mut self, font_scheme: &str) -> Format {
        self.font_scheme = font_scheme.to_string();
        self
    }

    /// Set the Format font family property.
    ///
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    ///
    /// # Arguments
    ///
    /// * `font_family` - The font family property.
    ///
    pub fn set_font_family(mut self, font_family: u8) -> Format {
        self.font_family = font_family;
        self
    }

    /// Set the Format font character set property.
    ///
    /// Set the font character. This function is implemented for completeness
    /// but is rarely used in practice.
    ///
    /// # Arguments
    ///
    /// * `font_charset` - The font character set property.
    ///
    pub fn set_font_charset(mut self, font_charset: u8) -> Format {
        self.font_charset = font_charset;
        self
    }

    /// Set the underline properties for a format.
    ///
    /// The difference between a normal underline and an "accounting" underline
    /// is that a normal underline only underlines the text/number in a cell
    /// whereas an accounting underline underlines the entire cell width.
    ///
    /// # Arguments
    ///
    /// * `underline` - The underline type defined by a [`XlsxUnderline`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting underline properties for a
    /// format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError, XlsxUnderline};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_underline(XlsxUnderline::None);
    ///     let format2 = Format::new().set_underline(XlsxUnderline::Single);
    ///     let format3 = Format::new().set_underline(XlsxUnderline::Double);
    ///     let format4 = Format::new().set_underline(XlsxUnderline::SingleAccounting);
    ///     let format5 = Format::new().set_underline(XlsxUnderline::DoubleAccounting);
    ///
    ///     worksheet.write_string(0, 0, "None",              &format1)?;
    ///     worksheet.write_string(1, 0, "Single",            &format2)?;
    ///     worksheet.write_string(2, 0, "Double",            &format3)?;
    ///     worksheet.write_string(3, 0, "Single Accounting", &format4)?;
    ///     worksheet.write_string(4, 0, "Double Accounting", &format5)?;
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
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_underline.png">
    ///
    pub fn set_underline(mut self, underline: XlsxUnderline) -> Format {
        self.underline = underline;
        self
    }

    /// Set the Format font strikeout property.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the text strikeout/strikethrough
    /// property for a format.
    ///
    /// ```
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file.
    /// #     let mut workbook = Workbook::new("formats.xlsx");
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_strikeout();
    ///
    ///     worksheet.write_string(0, 0, "Strikeout Text", &format)?;
    ///
    /// #     workbook.close()?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_font_strikeout.png">
    ///
    pub fn set_font_strikeout(mut self) -> Format {
        self.font_strikeout = true;
        self
    }

    /// Set the Format font super/subscript property.
    ///
    /// Note, this method is currently of limited use until the library
    /// implements multi-format "rich" strings.
    ///
    /// # Arguments
    ///
    /// * `font_script` - The font superscript or subscript property via a
    ///   [`XlsxScript`] enum.
    ///
    ///
    pub fn set_font_script(mut self, font_script: XlsxScript) -> Format {
        self.font_script = font_script;
        self
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs
// -----------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq)]
/// The XlsxColor enum defines an RGB color the can be used in rust_xlsxwriter
/// formatting.
///
/// You can use a small range of named colors or defined your own RGB color.
///
/// # Examples
///
/// The following example demonstrates using different XlsxColor enum values to
/// set the color of some text in a worksheet.
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("colors.xlsx");
///
///     let format1 = Format::new().set_font_color(XlsxColor::Red);
///     let format2 = Format::new().set_font_color(XlsxColor::Green);
///     let format3 = Format::new().set_font_color(XlsxColor::RGB(0x4F026A));
///     let format4 = Format::new().set_font_color(XlsxColor::RGB(0x73CC5F));
///     let format5 = Format::new().set_font_color(XlsxColor::RGB(0xFFACFF));
///     let format6 = Format::new().set_font_color(XlsxColor::RGB(0xCC7E16));
///
///     let worksheet = workbook.add_worksheet();
///     worksheet.write_string(0, 0, "Red", &format1)?;
///     worksheet.write_string(1, 0, "Green", &format2)?;
///     worksheet.write_string(2, 0, "#4F026A", &format3)?;
///     worksheet.write_string(3, 0, "#73CC5F", &format4)?;
///     worksheet.write_string(4, 0, "#FFACFF", &format5)?;
///     worksheet.write_string(5, 0, "#CC7E16", &format6)?;
///
/// #     workbook.close()?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/enum_xlsxcolor.png">
///
pub enum XlsxColor {
    /// A user defined RGB color in the range 0x000000 (black) to 0xFFFFFF
    /// (white). Any values outside this range will be ignored with a a warning.
    RGB(u32),

    /// The default/automatic color for an Excel property.
    Automatic,

    /// The color Black with a RGB value of 0x000000.
    Black,

    /// The color Blue with a RGB value of 0x0000FF.
    Blue,

    /// The color Brown with a RGB value of 0x800000.
    Brown,

    /// The color Cyan with a RGB value of 0x00FFFF.
    Cyan,

    /// The color Gray with a RGB value of 0x808080.
    Gray,

    /// The color Green with a RGB value of 0x008000.
    Green,

    /// The color Lime with a RGB value of 0x00FF00.
    Lime,

    /// The color Magenta with a RGB value of 0xFF00FF.
    Magenta,

    /// The color Navy with a RGB value of 0x000080.
    Navy,

    /// The color Orange with a RGB value of 0xFF6600.
    Orange,

    /// The color Pink with a RGB value of 0xFF00FF.
    Pink,

    /// The color Purple with a RGB value of 0x800080.
    Purple,

    /// The color Red with a RGB value of 0xFF0000.
    Red,

    /// The color Silver with a RGB value of 0xC0C0C0.
    Silver,

    /// The color White with a RGB value of 0xFFFFFF.
    White,

    /// The color Yellow with a RGB value of 0xFFFF00
    Yellow,
}

impl XlsxColor {
    // Get the u32 RGB value for a color.
    pub(crate) fn value(self) -> u32 {
        match self {
            XlsxColor::RGB(color) => color,
            XlsxColor::Automatic => 0xFFFFFFFF,
            XlsxColor::Black => 0x1000000,
            XlsxColor::Blue => 0x0000FF,
            XlsxColor::Brown => 0x800000,
            XlsxColor::Cyan => 0x00FFFF,
            XlsxColor::Gray => 0x808080,
            XlsxColor::Green => 0x008000,
            XlsxColor::Lime => 0x00FF00,
            XlsxColor::Magenta => 0xFF00FF,
            XlsxColor::Navy => 0x000080,
            XlsxColor::Orange => 0xFF6600,
            XlsxColor::Pink => 0xFF00FF,
            XlsxColor::Purple => 0x800080,
            XlsxColor::Red => 0xFF0000,
            XlsxColor::Silver => 0xC0C0C0,
            XlsxColor::White => 0xFFFFFF,
            XlsxColor::Yellow => 0xFFFF00,
        }
    }

    // Get the u32 RGB value for a color.
    pub(crate) fn hex_value(self) -> String {
        format!("{:06X}", self.value())
    }

    // Check if the RGB value is in the correct range.
    pub(crate) fn is_valid(self) -> bool {
        if let XlsxColor::RGB(color) = self {
            if color > 0xFFFFFF {
                eprintln!("RGB color must be in the the range 0x000000 - 0xFFFFFF");
                return false;
            }
        }

        true
    }
}

#[derive(Clone, Copy, PartialEq)]
/// The XlsxUnderline enum defines the font underline type in a format.
///
/// The difference between a normal underline and an "accounting" underline is
/// that a normal underline only underlines the text/number in a cell whereas an
/// accounting underline underlines the entire cell width.
///
/// # Examples
///
/// The following example demonstrates setting underline properties for a
/// format.
///
/// ```
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError, XlsxUnderline};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file.
/// #     let mut workbook = Workbook::new("formats.xlsx");
/// #     let worksheet = workbook.add_worksheet();
/// #
///     let format1 = Format::new().set_underline(XlsxUnderline::None);
///     let format2 = Format::new().set_underline(XlsxUnderline::Single);
///     let format3 = Format::new().set_underline(XlsxUnderline::Double);
///     let format4 = Format::new().set_underline(XlsxUnderline::SingleAccounting);
///     let format5 = Format::new().set_underline(XlsxUnderline::DoubleAccounting);
///
///     worksheet.write_string(0, 0, "None",              &format1)?;
///     worksheet.write_string(1, 0, "Single",            &format2)?;
///     worksheet.write_string(2, 0, "Double",            &format3)?;
///     worksheet.write_string(3, 0, "Single Accounting", &format4)?;
///     worksheet.write_string(4, 0, "Double Accounting", &format5)?;
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
/// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/format_set_underline.png">
///
pub enum XlsxUnderline {
    /// The default/automatic underline for an Excel font.
    None,

    /// A single underline under the text/number in a cell.
    Single,

    /// A double underline under the text/number in a cell.
    Double,

    /// A single accounting style underline under the entire cell.
    SingleAccounting,

    /// A double accounting style underline under the entire cell.
    DoubleAccounting,
}

#[derive(Clone, Copy, PartialEq)]
/// The XlsxScript enum defines the Format font superscript and subscript
/// properties.
///
pub enum XlsxScript {
    /// The default/automatic format for an Excel font.
    None,

    /// The cell text is superscripted.
    Superscript,

    /// The cell text is subscripted.
    Subscript,
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::XlsxColor;

    #[test]
    fn test_hex_value() {
        assert_eq!("FFFFFFFF", XlsxColor::Automatic.hex_value());
        assert_eq!("0000FF", XlsxColor::Blue.hex_value());
        assert_eq!("C0C0C0", XlsxColor::Silver.hex_value());
        assert_eq!("ABCDEF", XlsxColor::RGB(0xABCDEF).hex_value());
    }
}
