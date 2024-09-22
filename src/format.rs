// format - A module for representing Excel cell formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::{collections::HashMap, fmt, hash::Hash, sync::OnceLock};

use crate::Color;

/// The `Format` struct is used to define cell formatting for data in a
/// worksheet.
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
/// use rust_xlsxwriter::{Format, Workbook, FormatBorder, Color, XlsxError};
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
///     worksheet.write_string_with_format(0, 0, "Fonts", &format1)?;
///
///     let format2 = Format::new().set_font_name("Algerian").set_font_size(14);
///     worksheet.write_string_with_format(1, 0, "Fonts", &format2)?;
///
///     let format3 = Format::new().set_font_name("Comic Sans MS");
///     worksheet.write_string_with_format(2, 0, "Fonts", &format3)?;
///
///     let format4 = Format::new().set_font_name("Edwardian Script ITC");
///     worksheet.write_string_with_format(3, 0, "Fonts", &format4)?;
///
///     let format5 = Format::new().set_font_color(Color::Red);
///     worksheet.write_string_with_format(4, 0, "Font color", &format5)?;
///
///     let format6 = Format::new().set_background_color(Color::RGB(0xDAA520));
///     worksheet.write_string_with_format(5, 0, "Fills", &format6)?;
///
///     let format7 = Format::new().set_border(FormatBorder::Thin);
///     worksheet.write_string_with_format(6, 0, "Borders", &format7)?;
///
///     let format8 = Format::new().set_bold();
///     worksheet.write_string_with_format(7, 0, "Bold", &format8)?;
///
///     let format9 = Format::new().set_italic();
///     worksheet.write_string_with_format(8, 0, "Italic", &format9)?;
///
///     let format10 = Format::new().set_bold().set_italic();
///     worksheet.write_string_with_format(9, 0, "Bold and Italic", &format10)?;
///
///      workbook.save("formats.xlsx")?;
///
///      Ok(())
/// }
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
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
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
///         .set_font_color(Color::Red);
///
///     worksheet.write_string_with_format(0, 0, "Hello", &format)?;
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
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
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
///         .set_font_color(Color::Blue);
///
///     worksheet.write_string_with_format(0, 0, "Hello", &format1)?;
///     worksheet.write_string_with_format(1, 0, "Hello", &format2)?;
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
/// | Category        | Description           |  Method Name                             |
/// | :-------------- | :-------------------- |  :-------------------------------------- |
/// | **Number**      | Numeric format        |  [`Format::set_num_format()`]            |
/// | **Alignment**   | Horizontal align      |  [`Format::set_align()`]                 |
/// |                 | Vertical align        |  [`Format::set_align()`]                 |
/// |                 | Rotation              |  [`Format::set_rotation()`]              |
/// |                 | Text wrap             |  [`Format::set_text_wrap()`]             |
/// |                 | Indentation           |  [`Format::set_indent()`]                |
/// |                 | Reading direction     |  [`Format::set_reading_direction()`]     |
/// |                 | Shrink to fit         |  [`Format::set_shrink()`]                |
/// | **Font**        | Font type             |  [`Format::set_font_name()`]             |
/// |                 | Font size             |  [`Format::set_font_size()`]             |
/// |                 | Font color            |  [`Format::set_font_color()`]            |
/// |                 | Bold                  |  [`Format::set_bold()`]                  |
/// |                 | Italic                |  [`Format::set_italic()`]                |
/// |                 | Underline             |  [`Format::set_underline()`]             |
/// |                 | Strikethrough         |  [`Format::set_font_strikethrough()`]    |
/// |                 | Super/Subscript       |  [`Format::set_font_script()`]           |
/// | **Border**      | Cell border           |  [`Format::set_border()`]                |
/// |                 | Bottom border         |  [`Format::set_border_bottom()`]         |
/// |                 | Top border            |  [`Format::set_border_top()`]            |
/// |                 | Left border           |  [`Format::set_border_left()`]           |
/// |                 | Right border          |  [`Format::set_border_right()`]          |
/// |                 | Border color          |  [`Format::set_border_color()`]          |
/// |                 | Bottom color          |  [`Format::set_border_bottom_color()`]   |
/// |                 | Top color             |  [`Format::set_border_top_color()`]      |
/// |                 | Left color            |  [`Format::set_border_left_color()`]     |
/// |                 | Right color           |  [`Format::set_border_right_color()`]    |
/// |                 | Diagonal border       |  [`Format::set_border_diagonal()`]       |
/// |                 | Diagonal border color |  [`Format::set_border_diagonal_color()`] |
/// |                 | Diagonal border type  |  [`Format::set_border_diagonal_type()`]  |
/// | **Fill**        | Cell pattern          |  [`Format::set_pattern()`]               |
/// |                 | Background color      |  [`Format::set_background_color()`]      |
/// |                 | Foreground color      |  [`Format::set_foreground_color()`]      |
/// | **Protection**  | Unlock cells          |  [`Format::set_unlocked()`]              |
/// |                 | Hide formulas         |  [`Format::set_hidden()`]                |
///
/// # Format Colors
///
/// Format property colors are specified by using the [`Color`] enum with a Html
/// style RGB integer value or a limited number of defined colors:
///
/// ```
/// # // This code is available in examples/doc_enum_Color.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     let format1 = Format::new().set_font_color(Color::Red);
///     let format2 = Format::new().set_font_color(Color::Green);
///     let format3 = Format::new().set_font_color(Color::RGB(0x4F026A));
///     let format4 = Format::new().set_font_color(Color::RGB(0x73CC5F));
///     let format5 = Format::new().set_font_color(Color::RGB(0xFFACFF));
///     let format6 = Format::new().set_font_color(Color::RGB(0xCC7E16));
///
///     let worksheet = workbook.add_worksheet();
///     worksheet.write_string_with_format(0, 0, "Red", &format1)?;
///     worksheet.write_string_with_format(1, 0, "Green", &format2)?;
///     worksheet.write_string_with_format(2, 0, "#4F026A", &format3)?;
///     worksheet.write_string_with_format(3, 0, "#73CC5F", &format4)?;
///     worksheet.write_string_with_format(4, 0, "#FFACFF", &format5)?;
///     worksheet.write_string_with_format(5, 0, "#CC7E16", &format6)?;
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
///      worksheet.write_string(0, 0, "Hello")?;
///      worksheet.write_string_with_format(1, 0, "Hello", &format)?;
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
/// The [`Format::set_num_format()`] method is used to set the number format for
/// numbers used with
/// [`Worksheet::write_number_with_format()`](crate::Worksheet::write_number_with_format()):
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
///     worksheet.write_number_with_format(0, 0, 1234.56, &currency_format)?;
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
///     worksheet.write_number_with_format(0, 0, 1234.56, &currency_format)?;
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
///     worksheet.write_number_with_format(0, 0, 1234.56, &currency_format)?;
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
/// In order to replicate Excel's behavior all `rust_xlsxwriter` programs should
/// use US locale formatting which will then be rendered in the settings of your
/// host OS.
///
#[derive(Debug, Clone, Eq)]

pub struct Format {
    pub(crate) dxf_index: u32,
    pub(crate) font_index: u16,
    pub(crate) fill_index: u16,
    pub(crate) border_index: u16,
    pub(crate) has_font: bool,
    pub(crate) has_fill: bool,
    pub(crate) has_border: bool,

    // Properties listed in terms of the Excel dialog.

    // Number properties.
    pub(crate) num_format: String,
    pub(crate) num_format_index: u16,

    // Font properties.
    pub(crate) font: Font,

    // Alignment properties
    pub(crate) alignment: Alignment,

    // Border properties
    pub(crate) borders: Border,

    // Fill properties
    pub(crate) fill: Fill,

    // Protection properties.
    pub(crate) hidden: bool,
    pub(crate) locked: bool,

    // Non-UI properties.
    pub(crate) quote_prefix: bool,
    pub(crate) is_dxf_format: bool,
}

impl Hash for Format {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.font.hash(state);
        self.alignment.hash(state);
        self.borders.hash(state);
        self.fill.hash(state);

        self.num_format.hash(state);
        self.num_format_index.hash(state);
        self.hidden.hash(state);
        self.locked.hash(state);
        self.quote_prefix.hash(state);
    }
}

impl PartialEq for Format {
    fn eq(&self, other: &Self) -> bool {
        self.font == other.font
            && self.alignment == other.alignment
            && self.borders == other.borders
            && self.fill == other.fill
            && self.num_format == other.num_format
            && self.num_format_index == other.num_format_index
            && self.hidden == other.hidden
            && self.locked == other.locked
            && self.quote_prefix == other.quote_prefix
    }
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
    /// #
    /// # #[allow(unused_variables)]
    /// fn main() {
    ///
    ///     let format = Format::new();
    ///
    /// }
    /// ```
    pub fn new() -> Format {
        Format {
            dxf_index: 0,
            font_index: 0,
            fill_index: 0,
            border_index: 0,
            has_font: false,
            has_fill: false,
            has_border: false,

            font: Font::default(),
            alignment: Alignment::default(),
            fill: Fill::default(),
            borders: Border::default(),

            hidden: false,
            locked: true,
            num_format: String::new(),
            num_format_index: 0,
            quote_prefix: false,
            is_dxf_format: false,
        }
    }

    // -----------------------------------------------------------------------
    // Crate private methods.
    // -----------------------------------------------------------------------

    pub(crate) fn set_font_index(&mut self, font_index: u16, has_font: bool) {
        self.font_index = font_index;
        self.has_font = has_font;
    }

    // For DXF formats (Table and Conditional) check if the font has changed.
    pub(crate) fn has_dxf_font(&self) -> bool {
        self.font.bold
            || self.font.italic
            || self.font.underline != FormatUnderline::None
            || self.font.strikethrough
            || !self.font.color.is_auto_or_default()
    }

    // For DXF formats (Table and Conditional) check if the fill has changed.
    pub(crate) fn has_dxf_fill(&self) -> bool {
        self.fill.pattern != FormatPattern::None
            || !self.fill.background_color.is_auto_or_default()
            || !self.fill.foreground_color.is_auto_or_default()
    }

    pub(crate) fn set_fill_index(&mut self, fill_index: u16, has_fill: bool) {
        self.fill_index = fill_index;
        self.has_fill = has_fill;
    }

    pub(crate) fn set_border_index(&mut self, border_index: u16, has_border: bool) {
        self.border_index = border_index;
        self.has_border = has_border;
    }

    pub(crate) fn set_num_format_index_u16(&mut self, num_format_index: u16) {
        self.num_format_index = num_format_index;
    }

    // Check if the format has an alignment property set and requires a Styles
    // <alignment> element. This also handles a special case where Excel ignores
    // Bottom as a default.
    pub(crate) fn has_alignment(&self) -> bool {
        self.alignment.horizontal != FormatAlign::General
            || !(self.alignment.vertical == FormatAlign::General
                || self.alignment.vertical == FormatAlign::Bottom)
            || self.alignment.indent != 0
            || self.alignment.rotation != 0
            || self.alignment.text_wrap
            || self.alignment.shrink
            || self.alignment.reading_direction != 0
    }

    // Check if the format has an alignment property set and requires a Styles
    // "applyAlignment" attribute.
    pub(crate) fn apply_alignment(&self) -> bool {
        self.alignment.horizontal != FormatAlign::General
            || self.alignment.vertical != FormatAlign::General
            || self.alignment.indent != 0
            || self.alignment.rotation != 0
            || self.alignment.text_wrap
            || self.alignment.shrink
            || self.alignment.reading_direction != 0
    }

    // Check if the format has protection properties set.
    pub(crate) fn has_protection(&self) -> bool {
        self.hidden || !self.locked
    }

    // Check if the format is in the default/unmodified condition.
    pub(crate) fn is_default(&self) -> bool {
        static DEFAULT_STATE: OnceLock<Format> = OnceLock::new();
        let default_state = DEFAULT_STATE.get_or_init(Format::default);

        self == default_state
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
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in different locales]:
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// # Parameters
    ///
    /// - `num_format`: The number format property.
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
    /// #     // Set column width for clarity.
    /// #     worksheet.set_column_width(0, 20)?;
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
    ///     worksheet.write_number_with_format(0, 0, 1.23456,   &format1)?;
    ///     worksheet.write_number_with_format(1, 0, 1.23456  , &format2)?;
    ///     worksheet.write_number_with_format(2, 0, 1234.56,   &format3)?;
    ///     worksheet.write_number_with_format(3, 0, 1234.56,   &format4)?;
    ///     worksheet.write_number_with_format(4, 0, 44927.521, &format5)?;
    ///     worksheet.write_number_with_format(5, 0, 44927.521, &format6)?;
    ///     worksheet.write_number_with_format(6, 0, 44927.521, &format7)?;
    ///     worksheet.write_number_with_format(7, 0, 44927.521, &format8)?;
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
    pub fn set_num_format(mut self, num_format: impl Into<String>) -> Format {
        self.num_format = num_format.into();
        self
    }

    /// Set the number format for a Format using a legacy format index.
    ///
    /// This method is similar to [`Format::set_num_format()`] except that it
    /// uses an index to a limited number of Excel's built-in, and legacy,
    /// number formats.
    ///
    /// Unless you need to specifically access one of Excel's built-in number
    /// formats the [`Format::set_num_format()`] method is a better solution.
    /// This method is mainly included for backward compatibility and
    /// completeness.
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
    ///  - These formats can also be set via
    ///    [`Format::set_num_format()`].
    ///
    /// # Parameters
    ///
    /// - `num_format_index`: The index to one of the inbuilt formats shown in
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
    ///     worksheet.write_number_with_format(0, 0, 44927.521, &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_num_format_index.png">
    ///
    pub fn set_num_format_index(mut self, num_format_index: u8) -> Format {
        self.num_format_index = u16::from(num_format_index);

        // Also map the index to a format string. Mainly for DXF formats.
        let num_formats = HashMap::from([
            (1, "0"),
            (2, "0.00"),
            (3, "#,##0"),
            (4, "#,##0.00"),
            (5, "($#,##0_);($#,##0)"),
            (6, "($#,##0_);[Red]($#,##0)"),
            (7, "($#,##0.00_);($#,##0.00)"),
            (8, "($#,##0.00_);[Red]($#,##0.00)"),
            (9, "0%"),
            (10, "0.00%"),
            (11, "0.00E+00"),
            (12, "# ?/?"),
            (13, "# ??/??"),
            (14, "m/d/yy"),
            (15, "d-mmm-yy"),
            (16, "d-mmm"),
            (17, "mmm-yy"),
            (18, "h:mm AM/PM"),
            (19, "h:mm:ss AM/PM"),
            (20, "h:mm"),
            (21, "h:mm:ss"),
            (22, "m/d/yy h:mm"),
            (37, "(#,##0_);(#,##0)"),
            (38, "(#,##0_);[Red](#,##0)"),
            (39, "(#,##0.00_);(#,##0.00)"),
            (40, "(#,##0.00_);[Red](#,##0.00)"),
            (41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(_)"),
            (42, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(_)"),
            (43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(_)"),
            (44, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(_)"),
            (45, "mm:ss"),
            (46, "[h]:mm:ss"),
            (47, "mm:ss.0"),
            (48, "##0.0E+0"),
            (49, "@"),
        ]);

        if let Some(num_format) = num_formats.get(&num_format_index) {
            self.num_format = (*num_format).to_string();
        }

        self
    }

    /// Set the bold property for a Format font.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the bold property for a format.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_bold.rs
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
    ///     worksheet.write_string_with_format(0, 0, "Hello", &format)?;
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
        self.font.bold = true;
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
    ///     worksheet.write_string_with_format(0, 0, "Hello", &format)?;
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
        self.font.italic = true;
        self
    }

    /// Set the color property for the Format font.
    ///
    /// The `set_font_color()` method is used to set the font color in a cell.
    /// To set the color of a cell background use the `set_bg_color()` and
    /// `set_pattern()` methods.
    ///
    /// # Parameters
    ///
    /// - `color`: The font color property defined by a [`Color`] enum
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
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_font_color(Color::Red);
    ///
    ///     worksheet.write_string_with_format(0, 0, "Wheelbarrow", &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_font_color.png">
    ///
    pub fn set_font_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.font.color = color;
        }

        self
    }

    /// Set the Format font name property.
    ///
    /// Set the font for a cell format. Excel can only display fonts that are
    /// installed on the system that it is running on. Therefore it is generally
    /// best to use standard Excel fonts.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name property.
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
    ///     worksheet.write_string_with_format(0, 0, "Avenir Black Oblique", &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_font_name.png">
    ///
    pub fn set_font_name(mut self, font_name: impl Into<String>) -> Format {
        self.font.name = font_name.into();

        if self.font.name != "Calibri" {
            self.font.scheme = String::new();
        }

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
    /// # Parameters
    ///
    /// - `font_size`: The font size property.
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
    ///     worksheet.write_string_with_format(0, 0, "Font Size 30", &format)?;
    /// #
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
        self.font.size = font_size.into().to_string();
        self
    }

    /// Set the Format font scheme property.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    ///
    /// # Parameters
    ///
    /// - `font_scheme`: The font scheme property.
    ///
    pub fn set_font_scheme(mut self, font_scheme: impl Into<String>) -> Format {
        self.font.scheme = font_scheme.into();
        self
    }

    /// Set the Format font family property.
    ///
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    ///
    /// # Parameters
    ///
    /// - `font_family`: The font family property.
    ///
    pub fn set_font_family(mut self, font_family: u8) -> Format {
        self.font.family = font_family;
        self
    }

    /// Set the Format font character set property.
    ///
    /// Set the font character. This function is implemented for completeness
    /// but is rarely used in practice.
    ///
    /// # Parameters
    ///
    /// - `font_charset`: The font character set property.
    ///
    pub fn set_font_charset(mut self, font_charset: u8) -> Format {
        self.font.charset = font_charset;
        self
    }

    /// Set the underline properties for a format.
    ///
    /// The difference between a normal underline and an "accounting" underline
    /// is that a normal underline only underlines the text/number in a cell
    /// whereas an accounting underline underlines the entire cell width.
    ///
    /// # Parameters
    ///
    /// - `underline`: The underline type defined by a [`FormatUnderline`] enum
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
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError, FormatUnderline};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_underline(FormatUnderline::None);
    ///     let format2 = Format::new().set_underline(FormatUnderline::Single);
    ///     let format3 = Format::new().set_underline(FormatUnderline::Double);
    ///     let format4 = Format::new().set_underline(FormatUnderline::SingleAccounting);
    ///     let format5 = Format::new().set_underline(FormatUnderline::DoubleAccounting);
    ///
    ///     worksheet.write_string_with_format(0, 0, "None",              &format1)?;
    ///     worksheet.write_string_with_format(1, 0, "Single",            &format2)?;
    ///     worksheet.write_string_with_format(2, 0, "Double",            &format3)?;
    ///     worksheet.write_string_with_format(3, 0, "Single Accounting", &format4)?;
    ///     worksheet.write_string_with_format(4, 0, "Double Accounting", &format5)?;
    /// #
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
    pub fn set_underline(mut self, underline: FormatUnderline) -> Format {
        self.font.underline = underline;
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
    ///     worksheet.write_string_with_format(0, 0, "Strikethrough Text", &format)?;
    /// #
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
        self.font.strikethrough = true;
        self
    }

    /// Set the Format font super/subscript property.
    ///
    /// This feature is generally only useful when using a font in a "rich"
    /// string. See
    /// [`write_rich_string()`](crate::Worksheet::write_rich_string).
    ///
    /// # Parameters
    ///
    /// - `font_script`: The font superscript or subscript property via a
    ///   [`FormatScript`] enum.
    ///
    ///
    pub fn set_font_script(mut self, font_script: FormatScript) -> Format {
        self.font.script = font_script;
        self
    }

    /// Set the Format alignment properties.
    ///
    /// This method is used to set the horizontal and vertical data alignment
    /// within a cell.
    ///
    /// # Parameters
    ///
    /// - `align`: The vertical and or horizontal alignment direction as
    ///   defined by the [`FormatAlign`] enum.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting various cell alignment
    /// properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_align.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatAlign, XlsxError};
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
    ///         .set_align(FormatAlign::Center);
    ///
    ///     let format2 = Format::new()
    ///         .set_align(FormatAlign::Top)
    ///         .set_align(FormatAlign::Left);
    ///
    ///     let format3 = Format::new()
    ///         .set_align(FormatAlign::VerticalCenter)
    ///         .set_align(FormatAlign::Center);
    ///
    ///     let format4 = Format::new()
    ///         .set_align(FormatAlign::Bottom)
    ///         .set_align(FormatAlign::Right);
    ///
    ///     worksheet.write_string_with_format(0, 0, "Center", &format1)?;
    ///     worksheet.write_string_with_format(1, 0, "Top - Left", &format2)?;
    ///     worksheet.write_string_with_format(2, 0, "Center - Center", &format3)?;
    ///     worksheet.write_string_with_format(3, 0, "Bottom - Right", &format4)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_align.png">
    ///
    pub fn set_align(mut self, align: FormatAlign) -> Format {
        match align {
            FormatAlign::General => {
                self.alignment.horizontal = FormatAlign::General;
                self.alignment.vertical = FormatAlign::General;
            }
            FormatAlign::Center
            | FormatAlign::CenterAcross
            | FormatAlign::Distributed
            | FormatAlign::Fill
            | FormatAlign::Justify
            | FormatAlign::Left
            | FormatAlign::Right => {
                self.alignment.horizontal = align;
            }
            FormatAlign::Bottom
            | FormatAlign::Top
            | FormatAlign::VerticalCenter
            | FormatAlign::VerticalDistributed
            | FormatAlign::VerticalJustify => {
                self.alignment.vertical = align;
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
    /// [`Worksheet::set_row_height()`](crate::Worksheet::set_row_height()).
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
    ///     worksheet.write_string(0, 0, "Some text that isn't wrapped")?;
    ///     worksheet.write_string_with_format(1, 0, "Some text that is wrapped", &format1)?;
    ///     worksheet.write_string_with_format(2, 0, "Some text\nthat is\nwrapped\nat newlines", &format1)?;
    /// #
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
        self.alignment.text_wrap = true;
        self
    }

    /// Set the Format indent property.
    ///
    /// This method can be used to indent text in a cell.
    ///
    /// Indentation is a horizontal alignment property. It can be used in Excel
    /// in conjunction with the [`FormatAlign::Left`], [`FormatAlign::Right`]
    /// and [`FormatAlign::Distributed`] alignments. It will override any other
    /// horizontal properties that don't support indentation.
    ///
    /// # Parameters
    ///
    /// - `indent`: The indentation level for the cell.
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
    ///     worksheet.write_string(0, 0, "Indent 0")?;
    ///     worksheet.write_string_with_format(1, 0, "Indent 1", &format1)?;
    ///     worksheet.write_string_with_format(2, 0, "Indent 2", &format2)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_indent.png">
    ///
    pub fn set_indent(mut self, indent: u8) -> Format {
        self.alignment.indent = indent;
        self
    }

    /// Set the Format rotation property.
    ///
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270 to indicate text where the
    /// letters run from top to bottom.
    ///
    /// # Parameters
    ///
    /// - `rotation`: The rotation angle.
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
    ///     worksheet.write_string_with_format(0, 0, "Rust", &format1)?;
    ///     worksheet.write_string_with_format(1, 0, "Rust", &format2)?;
    ///     worksheet.write_string_with_format(2, 0, "Rust", &format3)?;
    /// #
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
            270 => self.alignment.rotation = 255,
            -90..=-1 => self.alignment.rotation = -rotation + 90,
            0..=90 => self.alignment.rotation = rotation,
            _ => eprintln!("Rotation outside range: -90 <= angle <= 90."),
        }

        self
    }

    /// Set the Format text reading order property.
    ///
    /// Set the text reading direction. This is useful when creating Arabic,
    /// Hebrew or other near or far eastern worksheets. It can be used in
    /// conjunction with the Worksheet
    /// [`set_right_to_left`](crate::Worksheet::set_right_to_left()) method
    /// which changes the cell display direction of the worksheet.
    ///
    /// # Parameters
    ///
    /// - `reading_direction`: The reading order property, should be 0, 1, or
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
    ///     worksheet.write_string(0, 0, "نص عربي / English text")?;
    ///     worksheet.write_string_with_format(1, 0, "نص عربي / English text", &format1)?;
    ///     worksheet.write_string_with_format(2, 0, "نص عربي / English text", &format2)?;
    /// #
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

        self.alignment.reading_direction = reading_direction;
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
    ///     worksheet.write_string_with_format(0, 0, "Shrink text to fit", &format1)?;
    /// #
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
        self.alignment.shrink = true;
        self
    }

    /// Set the Format pattern property.
    ///
    /// Set the pattern for a cell. The most commonly used pattern is
    /// [`FormatPattern::Solid`].
    ///
    /// To set the pattern colors see [`Format::set_background_color()`] and
    /// [`Format::set_foreground_color()`].
    ///
    /// # Parameters
    ///
    /// - `pattern`: The pattern property defined by a [`FormatPattern`] enum
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
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError, FormatPattern};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_background_color(Color::Green)
    ///         .set_pattern(FormatPattern::Solid);
    ///
    ///     let format2 = Format::new()
    ///         .set_background_color(Color::Yellow)
    ///         .set_foreground_color(Color::Red)
    ///         .set_pattern(FormatPattern::DarkVertical);
    ///
    ///     worksheet.write_string_with_format(0, 0, "Rust", &format1)?;
    ///     worksheet.write_blank(1, 0, &format2)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_pattern.png">
    ///
    pub fn set_pattern(mut self, pattern: FormatPattern) -> Format {
        self.fill.pattern = pattern;
        self
    }

    /// Set the Format pattern background color property.
    ///
    /// The `set_background_color` method can be used to set the background
    /// color of a pattern. Patterns are defined via the [`Format::set_pattern`]
    /// method. If a pattern hasn't been defined then a solid fill pattern is
    /// used as the default.
    ///
    /// # Parameters
    ///
    /// - `color`: The background color property defined by a [`Color`] enum
    ///   value or a type that can convert [`Into`] a [`Color`].
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the cell background color,
    /// with a default solid pattern.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_background_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_background_color(Color::Green);
    ///
    ///     worksheet.write_string_with_format(0, 0, "Rust", &format1)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_background_color.png">
    ///
    ///
    ///
    pub fn set_background_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.fill.background_color = color;
        }

        self
    }

    /// Set the Format pattern foreground color property.
    ///
    /// The `set_foreground_color` method can be used to set the
    /// foreground/pattern color of a pattern. Patterns are defined via the
    /// [`Format::set_pattern`] method.
    ///
    /// # Parameters
    ///
    /// - `color`: The foreground color property defined by a [`Color`] enum
    ///   value or a type that can convert [`Into`] a [`Color`].
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the foreground/pattern color.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_foreground_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError, FormatPattern};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_background_color(Color::Yellow)
    ///         .set_foreground_color(Color::Red)
    ///         .set_pattern(FormatPattern::DarkVertical);
    ///
    ///     worksheet.write_blank(0, 0, &format1)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_foreground_color.png">
    ///
    ///
    ///
    pub fn set_foreground_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.fill.foreground_color = color;
        }

        self
    }

    /// Set the Format border property.
    ///
    /// Set the cell border style. Individual border elements can be configured
    /// using the following methods with the same parameters:
    ///
    /// - [`Format::set_border_top()`]
    /// - [`Format::set_border_left()`]
    /// - [`Format::set_border_right()`]
    /// - [`Format::set_border_color()`]
    ///
    /// # Parameters
    ///
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting a cell border.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_border.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatBorder, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new().set_border(FormatBorder::Thin);
    ///     let format2 = Format::new().set_border(FormatBorder::Dotted);
    ///     let format3 = Format::new().set_border(FormatBorder::Double);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///     worksheet.write_blank(5, 1, &format3)?;
    /// #
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
    pub fn set_border(mut self, border: FormatBorder) -> Format {
        self.borders.top_style = border;
        self.borders.left_style = border;
        self.borders.right_style = border;
        self.borders.bottom_style = border;

        self
    }

    /// Set the Format border color property.
    ///
    /// Set the cell border color. Individual border elements can be configured
    /// using the following methods with the same parameters:
    ///
    /// - [`Format::set_border_top_color()`]
    /// - [`Format::set_border_left_color()`]
    /// - [`Format::set_border_right_color()`]
    /// - [`Format::set_border_bottom_color()`]
    ///
    /// # Parameters
    ///
    /// - `color`: The border color as defined by a [`Color`] enum value or
    ///   a type that can convert [`Into`] a [`Color`].
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting a cell border and color.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_border_color.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatBorder, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_border(FormatBorder::Thin)
    ///         .set_border_color(Color::Blue);
    ///
    ///     let format2 = Format::new()
    ///         .set_border(FormatBorder::Dotted)
    ///         .set_border_color(Color::Red);
    ///
    ///     let format3 = Format::new()
    ///         .set_border(FormatBorder::Double)
    ///         .set_border_color(Color::Green);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///     worksheet.write_blank(5, 1, &format3)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_border_color.png">
    ///
    pub fn set_border_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if !color.is_valid() {
            return self;
        }

        self.borders.top_color = color;
        self.borders.left_color = color;
        self.borders.right_color = color;
        self.borders.bottom_color = color;
        self
    }

    /// Set the cell top border style.
    ///
    /// See [`Format::set_border()`] for details.
    ///
    /// # Parameters
    ///
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    ///   value.
    ///
    pub fn set_border_top(mut self, border: FormatBorder) -> Format {
        self.borders.top_style = border;
        self
    }

    /// Set the cell top border color.
    ///
    /// See [`Format::set_border_color()`] for details.
    ///
    /// # Parameters
    ///
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_border_top_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.borders.top_color = color;
        }

        self
    }

    /// Set the cell bottom border style.
    ///
    /// See [`Format::set_border()`] for details.
    ///
    /// # Parameters
    ///
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    ///   value.
    ///
    pub fn set_border_bottom(mut self, border: FormatBorder) -> Format {
        self.borders.bottom_style = border;
        self
    }

    /// Set the cell bottom border color.
    ///
    /// See [`Format::set_border_color()`] for details.
    ///
    /// # Parameters
    ///
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_border_bottom_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.borders.bottom_color = color;
        }

        self
    }

    /// Set the cell left border style.
    ///
    /// See [`Format::set_border()`] for details.
    ///
    /// # Parameters
    ///
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    ///   value.
    ///
    pub fn set_border_left(mut self, border: FormatBorder) -> Format {
        self.borders.left_style = border;
        self
    }

    /// Set the cell left border color.
    ///
    /// See [`Format::set_border_color()`] for details.
    ///
    /// # Parameters
    ///
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_border_left_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.borders.left_color = color;
        }

        self
    }

    /// Set the cell right border style.
    ///
    /// See [`Format::set_border()`] for details.
    ///
    /// # Parameters
    ///
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    ///   value.
    ///
    pub fn set_border_right(mut self, border: FormatBorder) -> Format {
        self.borders.right_style = border;
        self
    }

    /// Set the cell right border color.
    ///
    /// See [`Format::set_border_color()`] for details.
    ///
    /// # Parameters
    ///
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_border_right_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.borders.right_color = color;
        }

        self
    }

    /// Set the Format border diagonal property.
    ///
    /// Set the cell border diagonal line style. This method should be used in
    /// conjunction with the [`Format::set_border_diagonal_type()`] method to
    /// set the diagonal type.
    ///
    /// # Parameters
    ///
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting cell diagonal borders.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_border_diagonal.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatBorder, Color, FormatDiagonalBorder, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format1 = Format::new()
    ///         .set_border_diagonal(FormatBorder::Thin)
    ///         .set_border_diagonal_type(FormatDiagonalBorder::BorderUp);
    ///
    ///     let format2 = Format::new()
    ///         .set_border_diagonal(FormatBorder::Thin)
    ///         .set_border_diagonal_type(FormatDiagonalBorder::BorderDown);
    ///
    ///     let format3 = Format::new()
    ///         .set_border_diagonal(FormatBorder::Thin)
    ///         .set_border_diagonal_type(FormatDiagonalBorder::BorderUpDown);
    ///
    ///     let format4 = Format::new()
    ///         .set_border_diagonal(FormatBorder::Thin)
    ///         .set_border_diagonal_type(FormatDiagonalBorder::BorderUpDown)
    ///         .set_border_diagonal_color(Color::Red);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    ///     worksheet.write_blank(5, 1, &format3)?;
    ///     worksheet.write_blank(7, 1, &format4)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/format_set_border_diagonal.png">
    ///
    pub fn set_border_diagonal(mut self, border: FormatBorder) -> Format {
        self.borders.diagonal_style = border;
        self
    }

    /// Set the cell diagonal border color.
    ///
    /// See [`Format::set_border_diagonal()`] for details.
    ///
    /// # Parameters
    ///
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_border_diagonal_color(mut self, color: impl Into<Color>) -> Format {
        let color = color.into();
        if color.is_valid() {
            self.borders.diagonal_color = color;
        }

        self
    }

    /// Set the cell diagonal border direction type.
    ///
    /// See [`Format::set_border_diagonal()`] for details.
    ///
    /// # Parameters
    ///
    /// - `border_type`: The diagonal border type as defined by a
    ///   [`FormatDiagonalBorder`] enum value.
    ///
    pub fn set_border_diagonal_type(mut self, border_type: FormatDiagonalBorder) -> Format {
        self.borders.diagonal_type = border_type;
        self
    }

    /// Set the hyperlink style.
    ///
    /// Set the hyperlink style for use with urls. This is usually set
    /// automatically when writing urls without a format applied.
    ///
    pub fn set_hyperlink(mut self) -> Format {
        self.font.is_hyperlink = true;
        self.font.color = Color::Theme(10, 0);
        self.font.underline = FormatUnderline::Single;
        self.font.scheme = String::new();

        self
    }

    /// Set the Format cell unlocked state.
    ///
    /// This method can be used to allow modification of a cell in a protected
    /// worksheet. In Excel, cell locking is turned on by default for all cells.
    /// However, it only has an effect if the worksheet has been protected using
    /// the [`Worksheet::protect()`](crate::Worksheet::protect) method.
    ///
    /// # Examples
    ///
    /// Example of cell locking and formula hiding in an Excel worksheet
    /// `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/app_worksheet_protection.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some format objects.
    ///     let unlocked = Format::new().set_unlocked();
    ///     let hidden = Format::new().set_hidden();
    ///
    ///     // Protect the worksheet to turn on cell locking.
    ///     worksheet.protect();
    ///
    ///     // Examples of cell locking and hiding.
    ///     worksheet.write_string(0, 0, "Cell B1 is locked. It cannot be edited.")?;
    ///     worksheet.write_formula(0, 1, "=1+2")?; // Locked by default.
    ///
    ///     worksheet.write_string(1, 0, "Cell B2 is unlocked. It can be edited.")?;
    ///     worksheet.write_formula_with_format(1, 1, "=1+2", &unlocked)?;
    ///
    ///     worksheet.write_string(2, 0, "Cell B3 is hidden. The formula isn't visible.")?;
    ///     worksheet.write_formula_with_format(2, 1, "=1+2", &hidden)?;
    /// #
    /// #     worksheet.write_string(4, 0, "Use Menu -> Review -> Unprotect Sheet")?;
    /// #     worksheet.write_string(5, 0, "to remove the worksheet protection.")?;
    /// #
    /// #     worksheet.autofit();
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet_protection.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_worksheet_protection.png">
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
    /// the worksheet has been protected using the
    /// [`Worksheet::protect()`](crate::Worksheet::protect) method.
    ///
    /// See the example above.
    ///
    pub fn set_hidden(mut self) -> Format {
        self.hidden = true;
        self
    }

    /// Set the `quote_prefix` property for a Format.
    ///
    /// Set the quote prefix property of a format to ensure a string is treated
    /// as a string after editing. This is the same as prefixing the string with
    /// a single quote in Excel. You don't need to add the quote to the string
    /// but you do need to add the format.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the quote prefix property for a
    /// format.
    ///
    /// ```
    /// # // This code is available in examples/doc_format_set_quote_prefix.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let format = Format::new().set_quote_prefix();
    ///
    ///     // If the "=Hello" string was edited in Excel it would change into an
    ///     // invalid formula and raise an error. The quote prefix adds a virtual quote
    ///     // to the start of the string and prevents this from happening.
    ///     worksheet.write_string_with_format(0, 0, "=Hello", &format)?;
    /// #
    /// #     workbook.save("formats.xlsx")?;
    /// #
    /// #     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/format_set_quote_prefix.png">
    ///
    pub fn set_quote_prefix(mut self) -> Format {
        self.quote_prefix = true;
        self
    }

    /// Unset the bold Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_bold()`].
    ///
    pub fn unset_bold(mut self) -> Format {
        self.font.bold = false;
        self
    }

    /// Unset the italic Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_italic()`].
    ///
    pub fn unset_italic(mut self) -> Format {
        self.font.italic = false;
        self
    }

    /// Unset the font strikethrough Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_font_strikethrough()`].
    ///
    pub fn unset_font_strikethrough(mut self) -> Format {
        self.font.strikethrough = false;
        self
    }

    /// Unset the text wrap Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_text_wrap()`].
    ///
    pub fn unset_text_wrap(mut self) -> Format {
        self.alignment.text_wrap = false;
        self
    }

    /// Unset the shrink Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_shrink()`].
    ///
    pub fn unset_shrink(mut self) -> Format {
        self.alignment.shrink = false;
        self
    }

    /// Set the locked Format property back to its default "on" state.
    ///
    /// The opposite of [`Format::set_unlocked()`].
    ///
    pub fn set_locked(mut self) -> Format {
        self.locked = true;
        self
    }

    /// Unset the hidden Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_hidden()`].
    ///
    pub fn unset_hidden(mut self) -> Format {
        self.hidden = false;
        self
    }

    /// Unset the hyperlink style.
    pub fn unset_hyperlink_style(mut self) -> Format {
        self.font.is_hyperlink = true;

        self
    }

    /// Unset the `quote_prefix` Format property back to its default "off" state.
    ///
    /// The opposite of [`Format::set_quote_prefix()`].
    ///
    pub fn unset_quote_prefix(mut self) -> Format {
        self.quote_prefix = false;
        self
    }
}

#[derive(Debug, Clone, Copy, Hash, PartialEq, Eq, Default)]
pub(crate) struct Alignment {
    pub(crate) horizontal: FormatAlign,
    pub(crate) vertical: FormatAlign,
    pub(crate) text_wrap: bool,
    pub(crate) justify_last: bool,
    pub(crate) rotation: i16,
    pub(crate) indent: u8,
    pub(crate) shrink: bool,
    pub(crate) reading_direction: u8,
}

#[derive(Debug, Clone, Hash, PartialEq, Eq, Default)]
pub(crate) struct Border {
    pub(crate) bottom_style: FormatBorder,
    pub(crate) top_style: FormatBorder,
    pub(crate) left_style: FormatBorder,
    pub(crate) right_style: FormatBorder,
    pub(crate) bottom_color: Color,
    pub(crate) top_color: Color,
    pub(crate) left_color: Color,
    pub(crate) right_color: Color,
    pub(crate) diagonal_style: FormatBorder,
    pub(crate) diagonal_color: Color,
    pub(crate) diagonal_type: FormatDiagonalBorder,
}

impl Border {
    // Check if the border is in the default/unmodified condition.
    pub(crate) fn is_default(&self) -> bool {
        static DEFAULT_STATE: OnceLock<Border> = OnceLock::new();
        let default_state = DEFAULT_STATE.get_or_init(Border::default);

        self == default_state
    }
}

#[derive(Debug, Clone, Hash, PartialEq, Eq)]
pub(crate) struct Font {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: FormatUnderline,
    pub(crate) name: String,
    pub(crate) size: String,
    pub(crate) color: Color,
    pub(crate) strikethrough: bool,
    pub(crate) script: FormatScript,
    pub(crate) family: u8,
    pub(crate) charset: u8,
    pub(crate) scheme: String,
    pub(crate) condense: bool,
    pub(crate) extend: bool,
    pub(crate) is_hyperlink: bool,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            name: "Calibri".to_string(),
            size: "11".to_string(),
            family: 2,
            scheme: "minor".to_string(),
            bold: Default::default(),
            italic: Default::default(),
            underline: FormatUnderline::default(),
            color: Color::default(),
            strikethrough: Default::default(),
            script: FormatScript::default(),
            charset: Default::default(),
            condense: Default::default(),
            extend: Default::default(),
            is_hyperlink: Default::default(),
        }
    }
}

#[derive(Debug, Clone, Hash, PartialEq, Eq, Default)]
pub(crate) struct Fill {
    pub(crate) foreground_color: Color,
    pub(crate) background_color: Color,
    pub(crate) pattern: FormatPattern,
}

// -----------------------------------------------------------------------
// Helper enums/structs/traits
// -----------------------------------------------------------------------

/// Convert a number format string to a [`Format`] object.
///
/// This From/Into trait provides a simple way to convert an Excel number format
/// string into a [`Format`] object. It is the equivalent of
/// `Format::new().set_num_format("string")`.
///
/// This is used as a syntactic shortcut for APIs that generally only require a
/// number format, like [`TableColumn`](crate::TableColumn).
///
impl From<&Format> for Format {
    fn from(value: &Format) -> Format {
        (*value).clone()
    }
}

/// Convert a number format string to a [`Format`] object.
impl From<&str> for Format {
    fn from(value: &str) -> Format {
        Format::new().set_num_format(value)
    }
}

/// Convert a number format string to a [`Format`] object.
impl From<&String> for Format {
    fn from(value: &String) -> Format {
        Format::new().set_num_format(value)
    }
}

/// Convert a number format string to a [`Format`] object.
impl From<String> for Format {
    fn from(value: String) -> Format {
        Format::new().set_num_format(value)
    }
}

/// The `FormatPattern` enum defines the Excel pattern types that can be added to
/// a [`Format`].
#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
pub enum FormatPattern {
    /// Automatic or Empty pattern.
    #[default]
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

impl fmt::Display for FormatPattern {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Solid => write!(f, "solid"),
            Self::DarkUp => write!(f, "darkUp"),
            Self::LightUp => write!(f, "lightUp"),
            Self::Gray125 => write!(f, "gray125"),
            Self::DarkGray => write!(f, "darkGray"),
            Self::DarkDown => write!(f, "darkDown"),
            Self::DarkGrid => write!(f, "darkGrid"),
            Self::Gray0625 => write!(f, "gray0625"),
            Self::LightGray => write!(f, "lightGray"),
            Self::LightDown => write!(f, "lightDown"),
            Self::LightGrid => write!(f, "lightGrid"),
            Self::MediumGray => write!(f, "mediumGray"),
            Self::DarkTrellis => write!(f, "darkTrellis"),
            Self::DarkVertical => write!(f, "darkVertical"),
            Self::LightTrellis => write!(f, "lightTrellis"),
            Self::LightVertical => write!(f, "lightVertical"),
            Self::DarkHorizontal => write!(f, "darkHorizontal"),
            Self::LightHorizontal => write!(f, "lightHorizontal"),
        }
    }
}

#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
/// The `FormatBorder` enum defines the Excel border types that can be added to
/// a [`Format`] pattern.
pub enum FormatBorder {
    /// No border.
    #[default]
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

impl fmt::Display for FormatBorder {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Thin => write!(f, "thin"),
            Self::Hair => write!(f, "hair"),
            Self::Thick => write!(f, "thick"),
            Self::Medium => write!(f, "medium"),
            Self::Dashed => write!(f, "dashed"),
            Self::Dotted => write!(f, "dotted"),
            Self::Double => write!(f, "double"),
            Self::DashDot => write!(f, "dashDot"),
            Self::DashDotDot => write!(f, "dashDotDot"),
            Self::MediumDashed => write!(f, "mediumDashed"),
            Self::SlantDashDot => write!(f, "slantDashDot"),
            Self::MediumDashDot => write!(f, "mediumDashDot"),
            Self::MediumDashDotDot => write!(f, "mediumDashDotDot"),
        }
    }
}

/// The `FormatDiagonalBorder` enum defines [`Format`] diagonal border types.
///
/// This is used with the [`Format::set_border_diagonal()`] method.
///
#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
pub enum FormatDiagonalBorder {
    /// The default/automatic format for an Excel font.
    #[default]
    None,

    /// Cell diagonal border from bottom left to top right.
    BorderUp,

    /// Cell diagonal border from top left to bottom right.
    BorderDown,

    /// Cell diagonal border in both directions.
    BorderUpDown,
}

/// The `FormatUnderline` enum defines the font underline type in a [`Format`].
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
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError, FormatUnderline};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
///     let format1 = Format::new().set_underline(FormatUnderline::None);
///     let format2 = Format::new().set_underline(FormatUnderline::Single);
///     let format3 = Format::new().set_underline(FormatUnderline::Double);
///     let format4 = Format::new().set_underline(FormatUnderline::SingleAccounting);
///     let format5 = Format::new().set_underline(FormatUnderline::DoubleAccounting);
///
///     worksheet.write_string_with_format(0, 0, "None",              &format1)?;
///     worksheet.write_string_with_format(1, 0, "Single",            &format2)?;
///     worksheet.write_string_with_format(2, 0, "Double",            &format3)?;
///     worksheet.write_string_with_format(3, 0, "Single Accounting", &format4)?;
///     worksheet.write_string_with_format(4, 0, "Double Accounting", &format5)?;
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
#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
pub enum FormatUnderline {
    /// The default/automatic underline for an Excel font.
    #[default]
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

/// The `FormatScript` enum defines the [`Format`] font superscript and subscript
/// properties.
///
#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
pub enum FormatScript {
    /// The default/automatic format for an Excel font.
    #[default]
    None,

    /// The cell text is superscripted.
    Superscript,

    /// The cell text is subscripted.
    Subscript,
}

#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
/// The `FormatAlign` enum defines the vertical and horizontal alignment properties
/// of a [`Format`].
///
pub enum FormatAlign {
    /// General/default alignment. The cell will use Excel's default for the
    /// data type, for example Left for text and Right for numbers.
    #[default]
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
