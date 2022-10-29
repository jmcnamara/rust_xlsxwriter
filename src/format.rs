// format - A module for representing Excel cell formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

#[derive(Clone)]
/// The Format struct is used to define cell formatting for data in a worksheet.
///
/// The properties of a cell that can be formatted include: fonts, colors,
/// patterns, borders, alignment and number formatting.
///
/// <img src="https://rustxlsxwriter.github.io/images/format_intro.png">
///
/// The output file above was created with the following code:
///
/// ```
/// # // This code is available in examples/doc_format_intro.rs
/// #
/// use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxColor, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet.
///     let worksheet = workbook.add_worksheet();
///
///     // Make the first column wider for clarity.
///     worksheet.set_column_width(0, 14)?;
///
///     // Create some sample formats to display
///     let format1 = Format::new().set_font_name("Arial");
///     worksheet.write_string(0, 0, "Fonts", &format1)?;
///
///     let format2 = Format::new().set_font_name("Algerian").set_font_size(14);
///     worksheet.write_string(1, 0, "Fonts", &format2)?;
///
///     let format3 = Format::new().set_font_name("Comic Sans MS");
///     worksheet.write_string(2, 0, "Fonts", &format3)?;
///
///     let format4 = Format::new().set_font_name("Edwardian Script ITC");
///     worksheet.write_string(3, 0, "Fonts", &format4)?;
///
///     let format5 = Format::new().set_font_color(XlsxColor::Red);
///     worksheet.write_string(4, 0, "Font color", &format5)?;
///
///     let format6 = Format::new().set_background_color(XlsxColor::RGB(0xDAA520));
///     worksheet.write_string(5, 0, "Fills", &format6)?;
///
///     let format7 = Format::new().set_border(XlsxBorder::Thin);
///     worksheet.write_string(6, 0, "Borders", &format7)?;
///
///     let format8 = Format::new().set_bold();
///     worksheet.write_string(7, 0, "Bold", &format8)?;
///
///     let format9 = Format::new().set_italic();
///     worksheet.write_string(8, 0, "Italic", &format9)?;
///
///     let format10 = Format::new().set_bold().set_italic();
///     worksheet.write_string(9, 0, "Bold and Italic", &format10)?;
///
///      workbook.save("formats.xlsx")?;
///
///      Ok(())
///  }
/// ```
///
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
/// # // This code is available in examples/doc_format_create.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
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
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img src="https://rustxlsxwriter.github.io/images/format_create.png">
///
/// Formats can be cloned in the usual way:
///
/// ```
/// # // This code is available in examples/doc_format_clone.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
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
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img src="https://rustxlsxwriter.github.io/images/format_clone.png">
///
///
/// # Format methods and Format properties
///
/// The following table shows the Excel format categories, in the order shown in
/// the Excel "Format Cell" dialog, and the equivalent `rust_xlsxwriter` Format
/// method:
///
/// | Category        | Description           |  Method Name                                                          |
/// | :-------------- | :-------------------- |  :------------------------------------------------------------------- |
/// | **Number**      | Numeric format        |  [`set_num_format()`](Format::set_num_format())                       |
/// | **Alignment**   | Horizontal align      |  [`set_align()`](Format::set_align())                                 |
/// |                 | Vertical align        |  [`set_align()`](Format::set_align())                                 |
/// |                 | Rotation              |  [`set_rotation()`](Format::set_rotation())                           |
/// |                 | Text wrap             |  [`set_text_wrap()`](Format::set_text_wrap())                         |
/// |                 | Indentation           |  [`set_indent()`](Format::set_indent())                               |
/// |                 | Reading direction     |  [`set_reading_direction()`](Format::set_reading_direction())         |
/// |                 | Shrink to fit         |  [`set_shrink()`](Format::set_shrink())                               |
/// | **Font**        | Font type             |  [`set_font_name()`](Format::set_font_name())                         |
/// |                 | Font size             |  [`set_font_size()`](Format::set_font_size())                         |
/// |                 | Font color            |  [`set_font_color()`](Format::set_font_color())                       |
/// |                 | Bold                  |  [`set_bold()`](Format::set_bold())                                   |
/// |                 | Italic                |  [`set_italic()`](Format::set_italic())                               |
/// |                 | Underline             |  [`set_underline()`](Format::set_underline())                         |
/// |                 | Strikethrough         |  [`set_font_strikethrough()`](Format::set_font_strikethrough())       |
/// |                 | Super/Subscript       |  [`set_font_script()`](Format::set_font_script())                     |
/// | **Border**      | Cell border           |  [`set_border()`](Format::set_border())                               |
/// |                 | Bottom border         |  [`set_border_bottom()`](Format::set_border_bottom())                 |
/// |                 | Top border            |  [`set_border_top()`](Format::set_border_top())                       |
/// |                 | Left border           |  [`set_border_left()`](Format::set_border_left())                     |
/// |                 | Right border          |  [`set_border_right()`](Format::set_border_right())                   |
/// |                 | Border color          |  [`set_border_color()`](Format::set_border_color())                   |
/// |                 | Bottom color          |  [`set_border_bottom_color()`](Format::set_border_bottom_color())     |
/// |                 | Top color             |  [`set_border_top_color()`](Format::set_border_top_color())           |
/// |                 | Left color            |  [`set_border_left_color()`](Format::set_border_left_color())         |
/// |                 | Right color           |  [`set_border_right_color()`](Format::set_border_right_color())       |
/// |                 | Diagonal border       |  [`set_border_diagonal()`](Format::set_border_diagonal())             |
/// |                 | Diagonal border color |  [`set_border_diagonal_color()`](Format::set_border_diagonal_color()) |
/// |                 | Diagonal border type  |  [`set_border_diagonal_type()`](Format::set_border_diagonal_type())   |
/// | **Fill**        | Cell pattern          |  [`set_pattern()`](Format::set_pattern())                             |
/// |                 | Background color      |  [`set_background_color()`](Format::set_background_color())           |
/// |                 | Foreground color      |  [`set_foreground_color()`](Format::set_foreground_color())           |
/// | **Protection**  | Unlock cells          |  [`set_unlocked()`](Format::set_unlocked())                           |
/// |                 | Hide formulas         |  [`set_hidden()`](Format::set_hidden())                               |
///
/// # Format Colors
///
/// Format property colors are specified by using the [`XlsxColor`] enum with a
/// Html style RGB integer value or a limited number of defined colors:
///
/// ```
/// # // This code is available in examples/doc_enum_xlsxcolor.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
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
/// #     workbook.save("colors.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img src="https://rustxlsxwriter.github.io/images/enum_xlsxcolor.png">
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
/// # // This code is available in examples/doc_format_default.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
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
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// <img src="https://rustxlsxwriter.github.io/images/format_default.png">
///
///
/// # Number Format Categories
///
/// The [`set_num_format()`](Format::set_num_format) method is used to set the
/// number format for numbers used with
/// [`write_number()`](super::Worksheet::write_number()):
///
/// ```
/// # // This code is available in examples/doc_format_currency1.rs
///
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet.
///     let worksheet = workbook.add_worksheet();
///
///     // Add a format.
///     let currency_format = Format::new().set_num_format("$#,##0.00");
///
///     worksheet.write_number(0, 0, 1234.56, &currency_format)?;
///
///     workbook.save("currency_format.xlsx")?;
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
/// - Time
/// - Percentage
/// - Fraction
/// - Scientific
/// - Text
/// - Custom
///
/// In the case of the example above the formatted output shows up as a Number
/// category:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency1.png">
///
/// If we wanted to have the number format display as a different category, such
/// as Currency, then would need to match the number format string used in the
/// code with the number format used by Excel. The easiest way to do this is to
/// open the Number Formatting dialog in Excel and set the required format:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency2.png">
///
/// Then, while still in the dialog, change to Custom. The format displayed is
/// the format used by Excel.
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency3.png">
///
/// If we put the format that we found (`"[$$-409]#,##0.00"`) into our previous
/// example and rerun it we will get a number format in the Currency category:
///
/// ```
/// # // This code is available in examples/doc_format_currency2.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #    // Create a new Excel file object.
/// #    let mut workbook = Workbook::new();
/// #
/// #    // Add a worksheet.
/// #    let worksheet = workbook.add_worksheet();
/// #
/// #    // Add a format.
///     let currency_format = Format::new().set_num_format("[$$-409]#,##0.00");
///
///     worksheet.write_number(0, 0, 1234.56, &currency_format)?;
///
/// #   workbook.save("currency_format.xlsx")?;
/// #
/// #   Ok(())
/// # }
/// ```
///
/// That give us the following updated output. Note that the number category is
/// now shown as Currency:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency4.png">
///
/// The same process can be used to find format strings for "Date" or
/// "Accountancy" formats.
///
/// # Number Formats in different locales
///
/// As shown in the previous section the `format.set_num_format()` method is
/// used to set the number format for `rust_xlsxwriter` formats. A common use
/// case is to set a number format with a "grouping/thousands" separator and a
/// "decimal" point:
///
/// ```
/// # // This code is available in examples/doc_format_locale.rs
/// #
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
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
///     workbook.save("number_format.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// In the US locale (and some others) where the number "grouping/thousands"
/// separator is `","` and the "decimal" point is `"."` which would be shown in
/// Excel as:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency5.png">
///
/// In other locales these values may be reversed or different. They are
/// generally set in the "Region" settings of Windows or Mac OS.  Excel handles
/// this by storing the number format in the file format in the US locale, in
/// this case `#,##0.00`, but renders it according to the regional settings of
/// the host OS. For example, here is the same, unmodified, output file shown
/// above in a German locale:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency6.png">
///
/// And here is the same file in a Russian locale. Note the use of a space as
/// the "grouping/thousands" separator:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency7.png">
///
/// In order to replicate Excel's behavior all XlsxWriter programs should use US
/// locale formatting which will then be rendered in the settings of your host
/// OS.
///

pub struct Format {
    pub(crate) xf_index: u32,
    pub(crate) font_index: u16,
    pub(crate) fill_index: u16,
    pub(crate) border_index: u16,
    pub(crate) has_font: bool,
    pub(crate) has_fill: bool,
    pub(crate) has_border: bool,

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
    pub(crate) font_strikethrough: bool,
    pub(crate) font_script: XlsxScript,
    pub(crate) font_family: u8,
    pub(crate) font_charset: u8,
    pub(crate) font_scheme: String,
    pub(crate) font_condense: bool,
    pub(crate) font_extend: bool,
    pub(crate) theme: u8,

    // Alignment properties.
    pub(crate) horizontal_align: XlsxAlign,
    pub(crate) vertical_align: XlsxAlign,
    pub(crate) text_wrap: bool,
    pub(crate) justify_last: bool,
    pub(crate) rotation: i16,
    pub(crate) indent: u8,
    pub(crate) shrink: bool,
    pub(crate) reading_direction: u8,

    // Border properties.
    pub(crate) border_bottom: XlsxBorder,
    pub(crate) border_top: XlsxBorder,
    pub(crate) border_left: XlsxBorder,
    pub(crate) border_right: XlsxBorder,
    pub(crate) border_bottom_color: XlsxColor,
    pub(crate) border_top_color: XlsxColor,
    pub(crate) border_left_color: XlsxColor,
    pub(crate) border_right_color: XlsxColor,
    pub(crate) border_diagonal: XlsxBorder,
    pub(crate) border_diagonal_color: XlsxColor,
    pub(crate) border_diagonal_type: XlsxDiagonalBorder,

    // Fill properties.
    pub(crate) foreground_color: XlsxColor,
    pub(crate) background_color: XlsxColor,
    pub(crate) pattern: XlsxPattern,

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
    /// # // This code is available in examples/doc_format_new.rs
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
            fill_index: 0,
            border_index: 0,
            has_font: false,
            has_fill: false,
            has_border: false,

            num_format: "".to_string(),
            num_format_index: 0,
            bold: false,
            italic: false,
            underline: XlsxUnderline::None,
            font_name: "Calibri".to_string(),
            font_size: 11.0,
            font_color: XlsxColor::Automatic,
            font_strikethrough: false,
            font_script: XlsxScript::None,
            font_family: 2,
            font_charset: 0,
            font_scheme: "".to_string(),
            font_condense: false,
            font_extend: false,
            theme: 0,
            hidden: false,
            locked: true,
            horizontal_align: XlsxAlign::General,
            vertical_align: XlsxAlign::General,
            text_wrap: false,
            justify_last: false,
            rotation: 0,
            foreground_color: XlsxColor::Automatic,
            background_color: XlsxColor::Automatic,
            pattern: XlsxPattern::None,
            border_bottom: XlsxBorder::None,
            border_top: XlsxBorder::None,
            border_left: XlsxBorder::None,
            border_right: XlsxBorder::None,
            border_diagonal: XlsxBorder::None,
            border_diagonal_type: XlsxDiagonalBorder::None,
            border_bottom_color: XlsxColor::Automatic,
            border_top_color: XlsxColor::Automatic,
            border_left_color: XlsxColor::Automatic,
            border_right_color: XlsxColor::Automatic,
            border_diagonal_color: XlsxColor::Automatic,
            indent: 0,
            shrink: false,
            reading_direction: 0,
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

    pub(crate) fn set_fill_index(&mut self, fill_index: u16, has_fill: bool) {
        self.fill_index = fill_index;
        self.has_fill = has_fill;
    }

    pub(crate) fn set_border_index(&mut self, border_index: u16, has_border: bool) {
        self.border_index = border_index;
        self.has_border = has_border;
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
            self.font_strikethrough,
            self.italic,
            self.theme,
            self.underline as u8,
        )
    }

    pub(crate) fn border_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}",
            self.border_bottom.value(),
            self.border_bottom_color.hex_value(),
            self.border_diagonal.value(),
            self.border_diagonal_color.hex_value(),
            self.border_diagonal_type as u8,
            self.border_left.value(),
            self.border_left_color.hex_value(),
            self.border_right.value(),
            self.border_right_color.hex_value(),
            self.border_top.value(),
            self.border_top_color.hex_value(),
        )
    }

    pub(crate) fn fill_key(&self) -> String {
        format!(
            "{}:{}:{}",
            self.pattern.value(),
            self.background_color.hex_value(),
            self.foreground_color.hex_value(),
        )
    }

    pub(crate) fn alignment_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}",
            self.indent,
            self.reading_direction,
            self.rotation,
            self.shrink,
            self.horizontal_align as u8,
            self.vertical_align as u8,
            self.justify_last,
            self.text_wrap,
        )
    }

    pub(crate) fn set_num_format_index_u16(&mut self, num_format_index: u16) {
        self.num_format_index = num_format_index;
    }

    // Check if the format has an alignment property set and requires a Styles
    // <alignment> element. This also handles a special case where Excel ignore
    // Bottom as a default.
    pub(crate) fn has_alignment(&self) -> bool {
        self.horizontal_align != XlsxAlign::General
            || !(self.vertical_align == XlsxAlign::General
                || self.vertical_align == XlsxAlign::Bottom)
            || self.indent != 0
            || self.rotation != 0
            || self.text_wrap
            || self.shrink
            || self.reading_direction != 0
    }

    // Check if the format has an alignment property set and requires a Styles
    // "applyAlignment" attribute.
    pub(crate) fn apply_alignment(&self) -> bool {
        self.horizontal_align != XlsxAlign::General
            || self.vertical_align != XlsxAlign::General
            || self.indent != 0
            || self.rotation != 0
            || self.text_wrap
            || self.shrink
            || self.reading_direction != 0
    }

    // Check if the format has protection properties set.
    pub(crate) fn has_protection(&self) -> bool {
        self.hidden || !self.locked
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
    /// See also [Number Format Categories] and [Number Formats in different
    /// locales].
    ///
    /// [Number Format Categories]: struct.Format.html#number-format-categories
    /// [Number Formats in different locales]:
    ///     struct.Format.html#number-formats-in-different-locales
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
    /// # // This code is available in examples/doc_format_set_num_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
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
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Note how the numbers above have been displayed by Excel in the output
    /// file according to the given number format:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_num_format.png">
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
    /// # // This code is available in examples/doc_format_set_num_format_index.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_num_format_index(15);
    ///
    ///     worksheet.write_number(0, 0, 44927.521, &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_num_format_index.png">
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
    /// # // This code is available in examples/doc_format_set_set_bold.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_bold();
    ///
    ///     worksheet.write_string(0, 0, "Hello", &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_bold.png">
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
    /// # // This code is available in examples/doc_format_set_italic.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_italic();
    ///
    ///     worksheet.write_string(0, 0, "Hello", &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_italic.png">
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
    /// # // This code is available in examples/doc_format_set_font_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_color(XlsxColor::Red);
    ///
    ///     worksheet.write_string(0, 0, "Wheelbarrow", &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_font_color.png">
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
    /// # // This code is available in examples/doc_format_set_font_name.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_name("Avenir Black Oblique");
    ///
    ///     worksheet.write_string(0, 0, "Avenir Black Oblique", &format)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_font_name.png">
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
    /// # // This code is available in examples/doc_format_set_font_size.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_size(30);
    ///
    ///     worksheet.write_string(0, 0, "Font Size 30", &format)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_font_size.png">
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
    /// # // This code is available in examples/doc_format_set_underline.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError, XlsxUnderline};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
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
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_underline.png">
    ///
    pub fn set_underline(mut self, underline: XlsxUnderline) -> Format {
        self.underline = underline;
        self
    }

    /// Set the Format font strikethrough property.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the text strikethrough
    /// property for a format.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_font_strikethrough.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_strikethrough();
    ///
    ///     worksheet.write_string(0, 0, "Strikethrough Text", &format)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_font_strikethrough.png">
    ///
    pub fn set_font_strikethrough(mut self) -> Format {
        self.font_strikethrough = true;
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

    /// Set the Format alignment properties.
    ///
    /// This method is used to set the horizontal and vertical data alignment
    /// within a cell.
    ///
    /// # Arguments
    ///
    /// * `align` - The vertical and or horizontal alignment direction as
    ///   defined by the [`XlsxAlign`] enum.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting various cell alignment
    /// properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_align.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxAlign, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Widen the rows/column for clarity.
    /// #     worksheet.set_row_height(1, 30)?;
    /// #     worksheet.set_row_height(2, 30)?;
    /// #     worksheet.set_row_height(3, 30)?;
    /// #     worksheet.set_column_width(0, 18)?;
    /// #
    /// #     // Create some alignment formats.
    ///     let format1 = Format::new()
    ///         .set_align(XlsxAlign::Center);
    ///
    ///     let format2 = Format::new()
    ///         .set_align(XlsxAlign::Top)
    ///         .set_align(XlsxAlign::Left);
    ///
    ///     let format3 = Format::new()
    ///         .set_align(XlsxAlign::VerticalCenter)
    ///         .set_align(XlsxAlign::Center);
    ///
    ///     let format4 = Format::new()
    ///         .set_align(XlsxAlign::Bottom)
    ///         .set_align(XlsxAlign::Right);
    ///
    ///     worksheet.write_string(0, 0, "Center", &format1)?;
    ///     worksheet.write_string(1, 0, "Top - Left", &format2)?;
    ///     worksheet.write_string(2, 0, "Center - Center", &format3)?;
    ///     worksheet.write_string(3, 0, "Bottom - Right", &format4)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_align.png">
    ///
    pub fn set_align(mut self, align: XlsxAlign) -> Format {
        match align {
            XlsxAlign::General => {
                self.horizontal_align = XlsxAlign::General;
                self.vertical_align = XlsxAlign::General;
            }
            XlsxAlign::Center
            | XlsxAlign::CenterAcross
            | XlsxAlign::Distributed
            | XlsxAlign::Fill
            | XlsxAlign::Justify
            | XlsxAlign::Left
            | XlsxAlign::Right => {
                self.horizontal_align = align;
            }
            XlsxAlign::Bottom
            | XlsxAlign::Top
            | XlsxAlign::VerticalCenter
            | XlsxAlign::VerticalDistributed
            | XlsxAlign::VerticalJustify => {
                self.vertical_align = align;
            }
        }

        self
    }

    /// Set the Format text wrap property.
    ///
    /// This method is used to turn on automatic text wrapping for text in a
    /// cell. If you wish to control where the string is wrapped you can add
    /// newlines to the text (see the example below).
    ///
    /// Excel generally adjusts the height of the cell to fit the wrapped text
    /// unless a explicit row height has be set via
    /// [`worksheet.set_row_height()`](super::Worksheet::set_row_height()).
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting an implicit (without newline)
    /// text wrap and a user defined text wrap (with newlines).
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_text_wrap.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_text_wrap();
    ///
    ///     worksheet.write_string_only(0, 0, "Some text that isn't wrapped")?;
    ///     worksheet.write_string(1, 0, "Some text that is wrapped", &format1)?;
    ///     worksheet.write_string(2, 0, "Some text\nthat is\nwrapped\nat newlines", &format1)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_text_wrap.png">
    ///
    pub fn set_text_wrap(mut self) -> Format {
        self.text_wrap = true;
        self
    }

    /// Set the Format indent property.
    ///
    /// This method can be used to indent text in a cell.
    ///
    /// Indentation is a horizontal alignment property. It can be used in Excel
    /// in conjunction with the [Left](XlsxAlign::Left),
    /// [Right](XlsxAlign::Right) and [Distributed](XlsxAlign::Distributed)
    /// alignments. It will override any other horizontal properties that don't
    /// support indentation.
    ///
    /// # Arguments
    ///
    /// * `indent` - The indentation level for the cell.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the indentation level for
    /// cell text.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_indent.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_indent(1);
    ///     let format2 = Format::new().set_indent(2);
    ///
    ///     worksheet.write_string_only(0, 0, "Indent 0")?;
    ///     worksheet.write_string(1, 0, "Indent 1", &format1)?;
    ///     worksheet.write_string(2, 0, "Indent 2", &format2)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_indent.png">
    ///
    pub fn set_indent(mut self, indent: u8) -> Format {
        self.indent = indent;
        self
    }

    /// Set the Format rotation property.
    ///
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270 to indicate text where the
    /// letters run from top to bottom.
    ///
    /// # Arguments
    ///
    /// * `rotation` - The rotation angle.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting text rotation for a cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_rotation.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Widen the rows/column for clarity.
    /// #     worksheet.set_row_height(0, 30)?;
    /// #     worksheet.set_row_height(1, 30)?;
    /// #     worksheet.set_row_height(2, 60)?;
    /// #
    /// #     // Create some alignment formats.
    ///     let format1 = Format::new().set_rotation(30);
    ///     let format2 = Format::new().set_rotation(-30);
    ///     let format3 = Format::new().set_rotation(270);
    ///
    ///     worksheet.write_string(0, 0, "Rust", &format1)?;
    ///     worksheet.write_string(1, 0, "Rust", &format2)?;
    ///     worksheet.write_string(2, 0, "Rust", &format3)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_rotation.png">
    ///
    pub fn set_rotation(mut self, rotation: i16) -> Format {
        match rotation {
            270 => self.rotation = 255,
            -90..=-1 => self.rotation = -rotation + 90,
            0..=90 => self.rotation = rotation,
            _ => eprintln!("Rotation rotation outside range: -90 <= angle <= 90."),
        }

        self
    }

    /// Set the Format text reading order property.
    ///
    /// Set the text reading direction. This is useful when creating Arabic,
    /// Hebrew or other near or far eastern worksheets. It can be used in
    /// conjunction with the Worksheet
    /// [`set_right_to_left`](super::Worksheet::set_right_to_left()) method
    /// which changes the cell display direction of the worksheet.
    ///
    /// # Arguments
    ///
    /// * `reading_direction` - The reading order property, should be 0, 1, or
    ///   2, where these values refer to:
    ///
    ///   0. The reading direction is determined heuristically by Excel
    ///      depending on the text. This is the default option.
    ///   1. The text is displayed Left-to-Right, like English.
    ///   2. The text is displayed Right-to-Left, like Hebrew or Arabic.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the text reading direction.
    /// This is useful when creating Arabic, Hebrew or other near or far eastern
    /// worksheets.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_reading_direction.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.set_column_width(0, 25)?;
    /// #
    ///     let format1 = Format::new().set_reading_direction(1);
    ///     let format2 = Format::new().set_reading_direction(2);
    ///
    ///     worksheet.write_string_only(0, 0, "نص عربي / English text")?;
    ///     worksheet.write_string(1, 0, "نص عربي / English text", &format1)?;
    ///     worksheet.write_string(2, 0, "نص عربي / English text", &format2)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_reading_direction.png">
    ///
    pub fn set_reading_direction(mut self, reading_direction: u8) -> Format {
        if reading_direction > 2 {
            eprintln!("Reading direction must be 0, 1 or 2.");
            return self;
        }

        self.reading_direction = reading_direction;
        self
    }

    /// Set the Format shrink property.
    ///
    /// This method can be used to shrink text so that it fits in a cell
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the text shrink format.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_shrink.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_shrink();
    ///
    ///     worksheet.write_string(0, 0, "Shrink text to fit", &format1)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_shrink.png">
    ///
    pub fn set_shrink(mut self) -> Format {
        self.shrink = true;
        self
    }

    /// Set the Format pattern property.
    ///
    /// Set the pattern for a cell. The most commonly used pattern is
    /// [`XlsxPattern::Solid`].
    ///
    /// To set the pattern colors see
    /// [`set_background_color()`](Format::set_background_color()) and
    /// [`set_foreground_color()`](Format::set_foreground_color()).
    ///
    /// # Arguments
    ///
    /// * `pattern` - The pattern property defined by a [`XlsxPattern`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the cell pattern (with
    /// colors).
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_pattern.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError, XlsxPattern};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_background_color(XlsxColor::Green)
    ///         .set_pattern(XlsxPattern::Solid);
    ///
    ///     let format2 = Format::new()
    ///         .set_background_color(XlsxColor::Yellow)
    ///         .set_foreground_color(XlsxColor::Red)
    ///         .set_pattern(XlsxPattern::DarkVertical);
    ///
    ///     worksheet.write_string(0, 0, "Rust", &format1)?;
    ///     worksheet.write_blank(1, 0, &format2)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_pattern.png">
    ///
    pub fn set_pattern(mut self, pattern: XlsxPattern) -> Format {
        self.pattern = pattern;
        self
    }

    /// Set the Format pattern background color property.
    ///
    /// The `set_background_color` method can be used to set the background
    /// color of a pattern. Patterns are defined via the
    /// [`set_pattern`](Format::set_pattern()) method. If a pattern hasn't been
    /// defined then a solid fill pattern is used as the default.
    ///
    /// # Arguments
    ///
    /// * `background_color` - The background color property defined by a
    ///   [`XlsxColor`] enum value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the cell background color,
    /// with a default solid pattern.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_background_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_background_color(XlsxColor::Green);
    ///
    ///     worksheet.write_string(0, 0, "Rust", &format1)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_background_color.png">
    ///
    ///
    ///
    pub fn set_background_color(mut self, background_color: XlsxColor) -> Format {
        if !background_color.is_valid() {
            return self;
        }

        self.background_color = background_color;
        self
    }

    /// Set the Format pattern foreground color property.
    ///
    /// The `set_foreground_color` method can be used to set the
    /// foreground/pattern color of a pattern. Patterns are defined via the
    /// [`set_pattern`](Format::set_pattern()) method.
    ///
    /// # Arguments
    ///
    /// * `foreground_color` - The foreground color property defined by a
    ///   [`XlsxColor`] enum value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the foreground/pattern color.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_foreground_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError, XlsxPattern};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_background_color(XlsxColor::Yellow)
    ///         .set_foreground_color(XlsxColor::Red)
    ///         .set_pattern(XlsxPattern::DarkVertical);
    ///
    ///     worksheet.write_blank(0, 0, &format1)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_foreground_color.png">
    ///
    ///
    ///
    pub fn set_foreground_color(mut self, foreground_color: XlsxColor) -> Format {
        if !foreground_color.is_valid() {
            return self;
        }

        self.foreground_color = foreground_color;
        self
    }

    /// Set the Format border property.
    ///
    /// Set the cell border style. Individual border elements can be configured
    /// using the following methods with the same parameters:
    ///
    /// - [`set_border_top()`](Format::set_border_top())
    /// - [`set_border_left()`](Format::set_border_left())
    /// - [`set_border_right()`](Format::set_border_right())
    /// - [`set_border_color()`](Format::set_border_color())
    ///
    /// # Arguments
    ///
    /// * `border` - The border property as defined by a [`XlsxBorder`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting a cell border.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_border.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_border(XlsxBorder::Thin);
    ///     let format2 = Format::new().set_border(XlsxBorder::Dotted);
    ///     let format3 = Format::new().set_border(XlsxBorder::Double);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///     worksheet.write_blank(5, 1, &format3)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_border.png">
    ///
    pub fn set_border(mut self, border: XlsxBorder) -> Format {
        self.border_top = border;
        self.border_left = border;
        self.border_right = border;
        self.border_bottom = border;

        self
    }

    /// Set the Format border color property.
    ///
    /// Set the cell border color. Individual border elements can be configured
    /// using the following methods with the same parameters:
    ///
    /// - [`set_border_top_color()`](Format::set_border_top_color())
    /// - [`set_border_left_color()`](Format::set_border_left_color())
    /// - [`set_border_right_color()`](Format::set_border_right_color())
    /// - [`set_border_color_color()`](Format::set_border_color())
    ///
    /// Note: it is only, currently possible to set a border around a single
    /// cell. To set a border around a range of cells you will need to create
    /// 4-8 individual border formats and apply them to the cells around the
    /// required border. A later version of this library will provide helper
    /// functions to do this more easily.
    ///
    /// # Arguments
    ///
    /// * `color` - The border color as defined by a [`XlsxColor`] enum value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting a cell border and color.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_border_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxColor, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_border(XlsxBorder::Thin)
    ///         .set_border_color(XlsxColor::Blue);
    ///
    ///     let format2 = Format::new()
    ///         .set_border(XlsxBorder::Dotted)
    ///         .set_border_color(XlsxColor::Red);
    ///
    ///     let format3 = Format::new()
    ///         .set_border(XlsxBorder::Double)
    ///         .set_border_color(XlsxColor::Green);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///     worksheet.write_blank(5, 1, &format3)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_border_color.png">
    ///
    pub fn set_border_color(mut self, color: XlsxColor) -> Format {
        if !color.is_valid() {
            return self;
        }

        self.border_top_color = color;
        self.border_left_color = color;
        self.border_right_color = color;
        self.border_bottom_color = color;
        self
    }

    /// Set the cell top border style. See
    /// [`set_border()`](Format::set_border()) for details.
    ///
    /// # Arguments
    ///
    /// * `border` - The border property as defined by a [`XlsxBorder`] enum
    ///   value.
    ///
    pub fn set_border_top(mut self, border: XlsxBorder) -> Format {
        self.border_top = border;
        self
    }

    /// Set the cell top border color. See
    /// [`set_border_color()`](Format::set_border_color()) for details.
    ///
    /// # Arguments
    ///
    /// * `color` - The border color as defined by a [`XlsxColor`] enum value.
    ///
    pub fn set_border_top_color(mut self, color: XlsxColor) -> Format {
        if !color.is_valid() {
            return self;
        }

        self.border_top_color = color;
        self
    }

    /// Set the cell bottom border style. See
    /// [`set_border()`](Format::set_border()) for details.
    ///
    /// # Arguments
    ///
    /// * `border` - The border property as defined by a [`XlsxBorder`] enum
    ///   value.
    ///
    pub fn set_border_bottom(mut self, border: XlsxBorder) -> Format {
        self.border_bottom = border;
        self
    }

    /// Set the cell bottom border color. See
    /// [`set_border_color()`](Format::set_border_color()) for details.
    ///
    /// # Arguments
    ///
    /// * `color` - The border color as defined by a [`XlsxColor`] enum value.
    ///
    pub fn set_border_bottom_color(mut self, color: XlsxColor) -> Format {
        if !color.is_valid() {
            return self;
        }

        self.border_bottom_color = color;
        self
    }

    /// Set the cell left border style. See
    /// [`set_border()`](Format::set_border()) for details.
    ///
    /// # Arguments
    ///
    /// * `border` - The border property as defined by a [`XlsxBorder`] enum
    ///   value.
    ///
    pub fn set_border_left(mut self, border: XlsxBorder) -> Format {
        self.border_left = border;
        self
    }

    /// Set the cell left border color. See
    /// [`set_border_color()`](Format::set_border_color()) for details.
    ///
    /// # Arguments
    ///
    /// * `color` - The border color as defined by a [`XlsxColor`] enum value.
    ///
    pub fn set_border_left_color(mut self, color: XlsxColor) -> Format {
        if !color.is_valid() {
            return self;
        }

        self.border_left_color = color;
        self
    }

    /// Set the cell right border style. See
    /// [`set_border()`](Format::set_border()) for details.
    ///
    /// # Arguments
    ///
    /// * `border` - The border property as defined by a [`XlsxBorder`] enum
    ///   value.
    ///
    pub fn set_border_right(mut self, border: XlsxBorder) -> Format {
        self.border_right = border;
        self
    }

    /// Set the cell right border color. See
    /// [`set_border_color()`](Format::set_border_color()) for details.
    ///
    /// # Arguments
    ///
    /// * `color` - The border color as defined by a [`XlsxColor`] enum value.
    ///
    pub fn set_border_right_color(mut self, color: XlsxColor) -> Format {
        if !color.is_valid() {
            return self;
        }

        self.border_right_color = color;
        self
    }

    /// Set the Format border diagonal property.
    ///
    /// Set the cell border diagonal line style. This method should be used in
    /// conjunction with the
    /// [`set_border_diagonal_type()`](Format::set_border_diagonal_type())
    /// method to set the diagonal type.
    ///
    /// # Arguments
    ///
    /// * `border` - The border property as defined by a [`XlsxBorder`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting cell diagonal borders.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_border_diagonal.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxColor, XlsxDiagonalBorder, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_border_diagonal(XlsxBorder::Thin)
    ///         .set_border_diagonal_type(XlsxDiagonalBorder::BorderUp);
    ///
    ///     let format2 = Format::new()
    ///         .set_border_diagonal(XlsxBorder::Thin)
    ///         .set_border_diagonal_type(XlsxDiagonalBorder::BorderDown);
    ///
    ///     let format3 = Format::new()
    ///         .set_border_diagonal(XlsxBorder::Thin)
    ///         .set_border_diagonal_type(XlsxDiagonalBorder::BorderUpDown);
    ///
    ///     let format4 = Format::new()
    ///         .set_border_diagonal(XlsxBorder::Thin)
    ///         .set_border_diagonal_type(XlsxDiagonalBorder::BorderUpDown)
    ///         .set_border_diagonal_color(XlsxColor::Red);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///     worksheet.write_blank(5, 1, &format3)?;
    ///     worksheet.write_blank(7, 1, &format4)?;
    ///
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_border_diagonal.png">
    ///
    pub fn set_border_diagonal(mut self, border: XlsxBorder) -> Format {
        self.border_diagonal = border;
        self
    }

    /// Set the cell diagonal border color. See
    /// [`set_border_diagonal()`](Format::set_border_diagonal()) for details.
    ///
    /// # Arguments
    ///
    /// * `color` - The border color as defined by a [`XlsxColor`] enum value.
    ///
    pub fn set_border_diagonal_color(mut self, color: XlsxColor) -> Format {
        if !color.is_valid() {
            return self;
        }

        self.border_diagonal_color = color;
        self
    }

    /// Set the cell diagonal border direction type. See
    /// [`set_border_diagonal()`](Format::set_border_diagonal()) for details.
    ///
    /// # Arguments
    ///
    /// * `border_type` - The diagonal border type as defined by a
    ///   [`XlsxDiagonalBorder`] enum value.
    ///
    pub fn set_border_diagonal_type(mut self, border_type: XlsxDiagonalBorder) -> Format {
        self.border_diagonal_type = border_type;
        self
    }

    /// Set the Format cell unlocked state.
    ///
    /// This method can be used to allow modification of a cell in a protected
    /// worksheet. In Excel, cell locking is turned on by default for all cells.
    /// However, it only has an effect if the worksheet has been protected using
    /// the `worksheet.protect()` method (not implemented yet).
    ///
    pub fn set_unlocked(mut self) -> Format {
        self.locked = false;
        self
    }

    /// Set the Format property to hide formulas in a cell.
    ///
    /// This method can be used to hide a formula while still displaying its
    /// result. This is generally used to hide complex calculations from end
    /// users who are only interested in the result. It only has an effect if
    /// the worksheet has been protected using the `worksheet.protect()` method
    /// (not implemented yet).
    ///
    pub fn set_hidden(mut self) -> Format {
        self.hidden = true;
        self
    }

    /// Unset the bold Format property back to its default "off" state.
    /// The opposite of [`set_bold()`](Format::set_bold()).
    pub fn unset_bold(mut self) -> Format {
        self.bold = false;
        self
    }

    /// Unset the italic Format property back to its default "off" state.
    /// The opposite of [`set_italic()`](Format::set_italic()).
    pub fn unset_italic(mut self) -> Format {
        self.italic = false;
        self
    }

    /// Unset the font strikethrough Format property back to its default "off" state.
    /// The opposite of [`set_font_strikethrough()`](Format::set_font_strikethrough()).
    pub fn unset_font_strikethrough(mut self) -> Format {
        self.font_strikethrough = false;
        self
    }

    /// Unset the text wrap Format property back to its default "off" state.
    /// The opposite of [`set_text_wrap()`](Format::set_text_wrap()).
    pub fn unset_text_wrap(mut self) -> Format {
        self.text_wrap = false;
        self
    }

    /// Unset the shrink Format property back to its default "off" state.
    /// The opposite of [`set_shrink()`](Format::set_shrink()).
    pub fn unset_shrink(mut self) -> Format {
        self.shrink = false;
        self
    }

    /// Set the locked Format property back to its default "on" state.
    /// The opposite of [`set_unlocked()`](Format::set_unlocked()).
    pub fn set_locked(mut self) -> Format {
        self.locked = true;
        self
    }

    /// Unset the hidden Format property back to its default "off" state.
    /// The opposite of [`set_hidden()`](Format::set_hidden()).
    pub fn unset_hidden(mut self) -> Format {
        self.hidden = false;
        self
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs
// -----------------------------------------------------------------------

#[derive(Clone, Copy, Eq, PartialEq)]
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
/// # // This code is available in examples/doc_enum_xlsxcolor.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
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
/// #     workbook.save("colors.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/enum_xlsxcolor.png">
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
            XlsxColor::Black => 0x000000,
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
                eprintln!("RGB color must be in the the range 0x000000 - 0xFFFFFF.");
                return false;
            }
        }

        true
    }
}

#[derive(Clone, Copy, Eq, PartialEq)]
/// The XlsxPattern enum defines the Excel patterns values that can be added to
/// a Format pattern.
pub enum XlsxPattern {
    /// Automatic or Empty pattern.
    None,

    /// Solid pattern.
    Solid,

    /// Medium gray pattern.
    MediumGray,

    /// Dark gray pattern.
    DarkGray,

    /// Light gray pattern.
    LightGray,

    /// Dark horizontal line pattern.
    DarkHorizontal,

    /// Dark vertical line pattern.
    DarkVertical,

    /// Dark diagonal stripe pattern.
    DarkDown,

    /// Reverse dark diagonal stripe pattern.
    DarkUp,

    /// Dark grid pattern.
    DarkGrid,

    /// Dark trellis pattern.
    DarkTrellis,

    /// Light horizontal Line pattern.
    LightHorizontal,

    /// Light vertical line pattern.
    LightVertical,

    /// Light diagonal stripe pattern.
    LightDown,

    /// Reverse light diagonal stripe pattern.
    LightUp,

    /// Light grid pattern.
    LightGrid,

    /// Light trellis pattern.
    LightTrellis,

    /// 12.5% gray pattern.
    Gray125,

    /// 6.25% gray pattern.
    Gray0625,
}

impl XlsxPattern {
    // Get the Excel string value for the pattern type.
    pub(crate) fn value(&self) -> &str {
        match self {
            XlsxPattern::None => "none",
            XlsxPattern::Solid => "solid",
            XlsxPattern::MediumGray => "mediumGray",
            XlsxPattern::DarkGray => "darkGray",
            XlsxPattern::LightGray => "lightGray",
            XlsxPattern::DarkHorizontal => "darkHorizontal",
            XlsxPattern::DarkVertical => "darkVertical",
            XlsxPattern::DarkDown => "darkDown",
            XlsxPattern::DarkUp => "darkUp",
            XlsxPattern::DarkGrid => "darkGrid",
            XlsxPattern::DarkTrellis => "darkTrellis",
            XlsxPattern::LightHorizontal => "lightHorizontal",
            XlsxPattern::LightVertical => "lightVertical",
            XlsxPattern::LightDown => "lightDown",
            XlsxPattern::LightUp => "lightUp",
            XlsxPattern::LightGrid => "lightGrid",
            XlsxPattern::LightTrellis => "lightTrellis",
            XlsxPattern::Gray125 => "gray125",
            XlsxPattern::Gray0625 => "gray0625",
        }
    }
}

#[derive(Clone, Copy, Eq, PartialEq)]
/// The XlsxBorder enum defines the Excel border types that can be added to
/// a Format pattern.
pub enum XlsxBorder {
    /// No border.
    None,

    /// Thin border style.
    Thin,

    /// Medium border style.
    Medium,

    /// Dashed border style.
    Dashed,

    /// Dotted border style.
    Dotted,

    /// Thick border style.
    Thick,

    /// Double border style.
    Double,

    /// Hair border style.
    Hair,

    /// Medium dashed border style.
    MediumDashed,

    /// Dash-dot border style.
    DashDot,

    /// Medium dash-dot border style.
    MediumDashDot,

    /// Dash-dot-dot border style.
    DashDotDot,

    /// Medium dash-dot-dot border style.
    MediumDashDotDot,

    /// Slant dash-dot border style.
    SlantDashDot,
}

impl XlsxBorder {
    // Get the Excel string value for the border type.
    pub(crate) fn value(&self) -> &str {
        match self {
            XlsxBorder::None => "none",
            XlsxBorder::Thin => "thin",
            XlsxBorder::Medium => "medium",
            XlsxBorder::Dashed => "dashed",
            XlsxBorder::Dotted => "dotted",
            XlsxBorder::Thick => "thick",
            XlsxBorder::Double => "double",
            XlsxBorder::Hair => "hair",
            XlsxBorder::MediumDashed => "mediumDashed",
            XlsxBorder::DashDot => "dashDot",
            XlsxBorder::MediumDashDot => "mediumDashDot",
            XlsxBorder::DashDotDot => "dashDotDot",
            XlsxBorder::MediumDashDotDot => "mediumDashDotDot",
            XlsxBorder::SlantDashDot => "slantDashDot",
        }
    }
}

#[derive(Clone, Copy, Eq, PartialEq)]
/// The XlsxDiagonalBorder enum defines the [`Format`] diagonal border types, see
/// [`set_border_diagonal()`](Format::set_border_diagonal()) .
///
pub enum XlsxDiagonalBorder {
    /// The default/automatic format for an Excel font.
    None,

    /// Cell diagonal border from bottom left to top right.
    BorderUp,

    /// Cell diagonal border from top left to bottom right.
    BorderDown,

    /// Cell diagonal border in both directions.
    BorderUpDown,
}

#[derive(Clone, Copy, Eq, PartialEq)]
/// The XlsxUnderline enum defines the font underline type in a [`Format`].
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
/// # // This code is available in examples/doc_format_set_align.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError, XlsxUnderline};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
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
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_set_underline.png">
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

#[derive(Clone, Copy, Eq, PartialEq)]
/// The XlsxScript enum defines the [`Format`] font superscript and subscript
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

#[derive(Clone, Copy, Eq, PartialEq)]
/// The XlsxAlign enum defines the vertical and horizontal alignment properties
/// for a [`Format`].
///
pub enum XlsxAlign {
    /// General/default alignment. The cell will use Excel's default for the
    /// data type, for example Left for text and Right for numbers.
    General,

    /// Align text to the left.
    Left,

    /// Center text horizontally.
    Center,

    /// Align text to the right.
    Right,

    /// Fill (repeat) the text horizontally across the cell.
    Fill,

    /// Aligns the text to the left and right of the cell, if the text exceeds
    /// the width of the cell.
    Justify,

    /// Center the text across the cell or cells that have this alignment. This
    /// is an older form of merged cells.
    CenterAcross,

    /// Distribute the words in the text evenly across the cell.
    Distributed,

    /// Align text to the top.
    Top,

    /// Align text to the bottom.
    Bottom,

    /// Center text vertically.
    VerticalCenter,

    /// Aligns the text to the top and bottom of the cell, if the text exceeds
    /// the height of the cell.
    VerticalJustify,

    /// Distribute the words in the text evenly from top to bottom in the cell.
    VerticalDistributed,
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::Format;
    use super::XlsxColor;

    #[test]
    fn test_hex_value() {
        assert_eq!("FFFFFFFF", XlsxColor::Automatic.hex_value());
        assert_eq!("000000", XlsxColor::Black.hex_value());
        assert_eq!("0000FF", XlsxColor::Blue.hex_value());
        assert_eq!("800000", XlsxColor::Brown.hex_value());
        assert_eq!("00FFFF", XlsxColor::Cyan.hex_value());
        assert_eq!("808080", XlsxColor::Gray.hex_value());
        assert_eq!("008000", XlsxColor::Green.hex_value());
        assert_eq!("00FF00", XlsxColor::Lime.hex_value());
        assert_eq!("FF00FF", XlsxColor::Magenta.hex_value());
        assert_eq!("000080", XlsxColor::Navy.hex_value());
        assert_eq!("FF6600", XlsxColor::Orange.hex_value());
        assert_eq!("FF00FF", XlsxColor::Pink.hex_value());
        assert_eq!("800080", XlsxColor::Purple.hex_value());
        assert_eq!("FF0000", XlsxColor::Red.hex_value());
        assert_eq!("C0C0C0", XlsxColor::Silver.hex_value());
        assert_eq!("FFFFFF", XlsxColor::White.hex_value());
        assert_eq!("FFFF00", XlsxColor::Yellow.hex_value());
        assert_eq!("ABCDEF", XlsxColor::RGB(0xABCDEF).hex_value());
    }

    #[test]
    fn test_unset() {
        let format1 = Format::default();
        let format2 = Format::new()
            .set_bold()
            .set_italic()
            .set_font_strikethrough()
            .set_text_wrap()
            .set_shrink()
            .set_unlocked()
            .set_hidden()
            .unset_bold()
            .unset_italic()
            .unset_font_strikethrough()
            .unset_text_wrap()
            .unset_shrink()
            .set_locked()
            .unset_hidden();

        assert_eq!(format1.format_key(), format2.format_key());
    }
}
