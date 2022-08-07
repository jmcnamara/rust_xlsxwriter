// format - A module for representing Excel cell formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::workbook::Workbook;

#[derive(Clone)]
pub struct Format {
    is_changed: bool,
    xf_index: u32,
    font_index: u16,
    has_font: bool,

    num_format: String,
    num_format_index: u16,
    bold: bool,
    italic: bool,
    underline: u8,
    font_name: String,
    font_size: u8,
    font_color: u32,
    font_strikeout: bool,
    font_outline: bool,
    font_shadow: bool,
    font_script: u8,
    font_family: u8,
    font_charset: u8,
    font_scheme: String,
    font_condense: bool,
    font_extend: bool,
    theme: u8,
    hidden: bool,
    locked: bool,
    text_horizontal_align: u8,
    text_wrap: bool,
    text_vertical_align: u8,
    text_justify_last: bool,
    rotation: u16,
    foreground_color: u32,
    background_color: u32,
    pattern: u8,
    bottom: u8,
    top: u8,
    left: u8,
    right: u8,
    diagonal_border: u8,
    diagonal_type: u8,
    bottom_color: u32,
    top_color: u32,
    left_color: u32,
    right_color: u32,
    diagonal_color: u32,
    indent: u8,
    shrink: bool,
    reading_order: u8,
}

impl Default for Format {
    fn default() -> Self {
        Self::new()
    }
}

impl Format {
    // Create a new Format struct.
    pub fn new() -> Format {
        Format {
            is_changed: false,
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
            font_color: 0x000000,
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
            foreground_color: 0x000000,
            background_color: 0x000000,
            pattern: 0,
            bottom: 0,
            top: 0,
            left: 0,
            right: 0,
            diagonal_border: 0,
            diagonal_type: 0,
            bottom_color: 0x000000,
            top_color: 0x000000,
            left_color: 0x000000,
            right_color: 0x000000,
            diagonal_color: 0x000000,
            indent: 0,
            shrink: false,
            reading_order: 0,
        }
    }

    // -----------------------------------------------------------------------
    // Property getters.
    // -----------------------------------------------------------------------

    pub(crate) fn xf_index(&self) -> u32 {
        self.xf_index
    }

    pub(crate) fn has_font(&self) -> bool {
        self.has_font
    }

    pub(crate) fn get_font_index(&self) -> u16 {
        self.font_index
    }

    pub(crate) fn num_format(&self) -> &String {
        &self.num_format
    }

    pub(crate) fn num_format_index(&self) -> u16 {
        self.num_format_index
    }

    pub(crate) fn bold(&self) -> bool {
        self.bold
    }

    pub(crate) fn italic(&self) -> bool {
        self.italic
    }

    // -----------------------------------------------------------------------
    // Crate private methods.
    // -----------------------------------------------------------------------

    pub(crate) fn set_xf_index(&mut self, index: u32) {
        self.xf_index = index;
        self.is_changed = false;
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
            self.font_color,
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
            self.bottom_color,
            self.diagonal_border,
            self.diagonal_color,
            self.diagonal_type,
            self.left,
            self.left_color,
            self.right,
            self.right_color,
            self.top,
            self.top_color
        )
    }

    pub fn get_fill_key(&self) -> String {
        format!(
            "{}:{}:{}",
            self.background_color, self.foreground_color, self.pattern
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

    pub fn register_with(mut self, workbook: &mut Workbook) -> Format {
        workbook.register_format(&mut self);
        self
    }

    pub fn set_num_format(mut self, num_format: &str) -> Format {
        if self.num_format != num_format {
            self.num_format = num_format.to_string();
            self.is_changed = true;
        }
        self
    }

    pub fn set_num_format_index(mut self, num_format_index: u8) -> Format {
        let num_format_index = num_format_index as u16;
        if self.num_format_index != num_format_index {
            self.num_format_index = num_format_index;
            self.is_changed = true;
        }
        self
    }

    pub fn set_bold(mut self) -> Format {
        self.bold = true;
        self.is_changed = true;

        self
    }

    pub fn set_italic(mut self) -> Format {
        self.italic = true;
        self.is_changed = true;

        self
    }
}
