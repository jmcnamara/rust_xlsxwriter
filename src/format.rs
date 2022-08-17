// format - A module for representing Excel cell formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//#![warn(missing_docs)]
//use crate::workbook::Workbook;

#[derive(Clone)]
/// Format docs: TODO
pub struct Format {
    pub(crate) xf_index: u32,
    pub(crate) font_index: u16,
    pub(crate) has_font: bool,

    pub(crate) num_format: String,
    pub(crate) num_format_index: u16,
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: u8,
    pub(crate) font_name: String,
    pub(crate) font_size: u8,
    pub(crate) font_color: XlsxColor,
    pub(crate) font_strikeout: bool,
    pub(crate) font_outline: bool,
    pub(crate) font_shadow: bool,
    pub(crate) font_script: u8,
    pub(crate) font_family: u8,
    pub(crate) font_charset: u8,
    pub(crate) font_scheme: String,
    pub(crate) font_condense: bool,
    pub(crate) font_extend: bool,
    pub(crate) theme: u8,
    pub(crate) hidden: bool,
    pub(crate) locked: bool,
    pub(crate) text_horizontal_align: u8,
    pub(crate) text_wrap: bool,
    pub(crate) text_vertical_align: u8,
    pub(crate) text_justify_last: bool,
    pub(crate) rotation: u16,
    pub(crate) foreground_color: XlsxColor,
    pub(crate) background_color: XlsxColor,
    pub(crate) pattern: u8,
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
    pub(crate) indent: u8,
    pub(crate) shrink: bool,
    pub(crate) reading_order: u8,
}

impl Default for Format {
    fn default() -> Self {
        Self::new()
    }
}

impl Format {
    ///  Create a new Format struct.
    pub fn new() -> Format {
        Format {
            xf_index: 0,
            font_index: 0,
            has_font: false,

            num_format: "".to_string(),
            num_format_index: 0,
            bold: false,
            italic: false,
            underline: 0,
            font_name: "Calibri".to_string(),
            font_size: 11,
            font_color: XlsxColor::Automatic,
            font_strikeout: false,
            font_outline: false,
            font_shadow: false,
            font_script: 0,
            font_family: 2,
            font_charset: 0,
            font_scheme: "minor".to_string(),
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

    pub(crate) fn get_format_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}",
            self.get_alignment_key(),
            self.get_border_key(),
            self.get_fill_key(),
            self.get_font_key(),
            self.hidden,
            self.locked,
            self.num_format,
            self.num_format_index,
        )
    }

    pub(crate) fn get_font_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}",
            self.bold,
            self.font_charset,
            self.font_color.value(),
            self.font_condense,
            self.font_extend,
            self.font_family,
            self.font_name,
            self.font_outline,
            self.font_scheme,
            self.font_script,
            self.font_shadow,
            self.font_size,
            self.font_strikeout,
            self.italic,
            self.theme,
            self.underline,
        )
    }

    pub(crate) fn get_border_key(&self) -> String {
        format!(
            "{}:{}:{}:{}:{}:{}:{}:{}:{}:{}:{}",
            self.bottom,
            self.bottom_color.value(),
            self.diagonal_border,
            self.diagonal_color.value(),
            self.diagonal_type,
            self.left,
            self.left_color.value(),
            self.right,
            self.right_color.value(),
            self.top,
            self.top_color.value(),
        )
    }

    pub(crate) fn get_fill_key(&self) -> String {
        format!(
            "{}:{}:{}",
            self.background_color.value(),
            self.foreground_color.value(),
            self.pattern,
        )
    }

    pub(crate) fn get_alignment_key(&self) -> String {
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

    /// Temp function. Remove or document later. TODO.
    //pub fn register_with(mut self, workbook: &mut Workbook) -> Format {
    //    workbook.register_format(&mut self);
    //    self
    //}

    pub fn set_num_format(mut self, num_format: &str) -> Format {
        if self.num_format != num_format {
            self.num_format = num_format.to_string();
        }
        self
    }

    pub fn set_num_format_index(mut self, num_format_index: u8) -> Format {
        let num_format_index = num_format_index as u16;
        if self.num_format_index != num_format_index {
            self.num_format_index = num_format_index;
        }
        self
    }

    pub fn set_bold(mut self) -> Format {
        self.bold = true;
        self
    }

    pub fn set_italic(mut self) -> Format {
        self.italic = true;
        self
    }

    pub fn set_font_color(mut self, font_color: XlsxColor) -> Format {
        if !font_color.is_valid() {
            return self;
        }

        if self.font_color != font_color {
            self.font_color = font_color;
        }

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

    // Check if
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
