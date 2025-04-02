// styles - A module for creating the Excel styles.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;
use std::mem;

use crate::format::Format;
use crate::xmlwriter::{
    xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag, xml_start_tag,
    xml_start_tag_only,
};
use crate::{
    Alignment, Border, Color, Fill, Font, FormatAlign, FormatBorder, FormatDiagonalBorder,
    FormatPattern, FormatScript, FormatUnderline,
};

pub struct Styles<'a> {
    pub(crate) writer: Cursor<Vec<u8>>,
    xf_formats: &'a Vec<Format>,
    dxf_formats: &'a Vec<Format>,
    font_count: u16,
    fill_count: u16,
    border_count: u16,
    num_formats: Vec<String>,
    has_hyperlink_style: bool,
    has_comments: bool,
    is_rich_string_style: bool,
    hyperlink_font_id: u16,
}

impl<'a> Styles<'a> {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Styles struct.
    #[allow(clippy::similar_names)]
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn new(
        xf_formats: &'a Vec<Format>,
        dxf_formats: &'a Vec<Format>,
        font_count: u16,
        fill_count: u16,
        border_count: u16,
        num_formats: Vec<String>,
        has_hyperlink_style: bool,
        has_comments: bool,
        is_rich_string_style: bool,
    ) -> Styles<'a> {
        let writer = Cursor::new(Vec::with_capacity(2048));

        Styles {
            writer,
            xf_formats,
            dxf_formats,
            font_count,
            fill_count,
            border_count,
            num_formats,
            has_hyperlink_style,
            has_comments,
            is_rich_string_style,
            hyperlink_font_id: 0,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the styleSheet element.
        self.write_style_sheet();

        // Write the numFmts element.
        self.write_num_fmts();

        // Write the fonts element.
        self.write_fonts();

        // Write the fills element.
        self.write_fills();

        // Write the borders element.
        self.write_borders();

        // Write the cellStyleXfs element.
        self.write_cell_style_xfs();

        // Write the cellXfs element.
        self.write_cell_xfs();

        // Write the cellStyles element.
        self.write_cell_styles();

        // Write the dxfs element.
        self.write_dxfs();

        // Write the tableStyles element.
        self.write_table_styles();

        // Close the styleSheet tag.
        xml_end_tag(&mut self.writer, "styleSheet");
    }

    // Write the <styleSheet> element.
    fn write_style_sheet(&mut self) {
        let attributes = [(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )];

        xml_start_tag(&mut self.writer, "styleSheet", &attributes);
    }

    // Write the <fonts> element.
    fn write_fonts(&mut self) {
        let mut count = self.font_count;
        if self.has_comments {
            count += 1;
        }

        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "fonts", &attributes);

        // Write the cell font elements.
        let mut font_id = 0;
        for xf_format in self.xf_formats {
            // Write the font element.
            if xf_format.has_font {
                self.write_font(&xf_format.font, false);

                if xf_format.font.is_hyperlink {
                    self.hyperlink_font_id = font_id;
                }

                font_id += 1;
            }
        }

        // Add the additional font for cell comments/notes. This isn't currently
        // configurable and probably won't be.
        if self.has_comments {
            self.write_comment_font();
        }

        xml_end_tag(&mut self.writer, "fonts");
    }

    // Write the <font> element.
    pub(crate) fn write_font(&mut self, font: &Font, dxf_format: bool) {
        if self.is_rich_string_style {
            xml_start_tag_only(&mut self.writer, "rPr");
        } else {
            xml_start_tag_only(&mut self.writer, "font");
        }

        if font.bold {
            xml_empty_tag_only(&mut self.writer, "b");
        }

        if font.italic {
            xml_empty_tag_only(&mut self.writer, "i");
        }

        if font.strikethrough {
            xml_empty_tag_only(&mut self.writer, "strike");
        }

        if font.underline != FormatUnderline::None {
            self.write_font_underline(font);
        }

        if font.script != FormatScript::None {
            self.write_vert_align(font);
        }
        // Write the <sz> element.
        if !dxf_format {
            self.write_font_size(font);
        }

        // Write the <color> element.
        self.write_font_color(font, dxf_format);

        if !dxf_format {
            // Write the name element.
            self.write_font_name(font);

            // Write the family element.
            if font.family > 0 {
                self.write_font_family(font);
            }

            // Write the charset element.
            if font.charset > 0 {
                self.write_font_charset(font);
            }

            // Write the scheme element.
            self.write_font_scheme(font);
        }

        if self.is_rich_string_style {
            xml_end_tag(&mut self.writer, "rPr");
        } else {
            xml_end_tag(&mut self.writer, "font");
        }
    }

    // Write the <sz> element.
    fn write_font_size(&mut self, font: &Font) {
        let attributes = [("val", font.size.as_str())];

        xml_empty_tag(&mut self.writer, "sz", &attributes);
    }

    // Write the <color> element.
    fn write_font_color(&mut self, font: &Font, dxf_format: bool) {
        let mut attributes = vec![];

        match font.color {
            Color::Automatic => {
                // The color element is omitted for an Automatic color.
            }
            Color::Default => {
                if !dxf_format {
                    attributes.push(("theme", "1".to_string()));
                    xml_empty_tag(&mut self.writer, "color", &attributes);
                }
            }
            _ => {
                attributes.append(&mut font.color.attributes());
                xml_empty_tag(&mut self.writer, "color", &attributes);
            }
        }
    }

    // Write the <name> element.
    fn write_font_name(&mut self, font: &Font) {
        let attributes = [("val", font.name.as_ref())];

        if self.is_rich_string_style {
            xml_empty_tag(&mut self.writer, "rFont", &attributes);
        } else {
            xml_empty_tag(&mut self.writer, "name", &attributes);
        }
    }

    // Write the <family> element.
    fn write_font_family(&mut self, font: &Font) {
        let attributes = [("val", font.family.to_string())];

        xml_empty_tag(&mut self.writer, "family", &attributes);
    }

    // Write the <charset> element.
    fn write_font_charset(&mut self, font: &Font) {
        let attributes = [("val", font.charset.to_string())];

        xml_empty_tag(&mut self.writer, "charset", &attributes);
    }

    // Write the <scheme> element.
    fn write_font_scheme(&mut self, font: &Font) {
        let mut attributes = vec![];

        if !font.scheme.is_empty() {
            attributes.push(("val", font.scheme.to_string()));
            xml_empty_tag(&mut self.writer, "scheme", &attributes);
        }
    }

    // Write the <u> underline element.
    fn write_font_underline(&mut self, font: &Font) {
        let mut attributes = vec![];

        match font.underline {
            FormatUnderline::Double => {
                attributes.push(("val", "double".to_string()));
            }
            FormatUnderline::SingleAccounting => {
                attributes.push(("val", "singleAccounting".to_string()));
            }
            FormatUnderline::DoubleAccounting => {
                attributes.push(("val", "doubleAccounting".to_string()));
            }
            _ => {}
        }

        xml_empty_tag(&mut self.writer, "u", &attributes);
    }

    // Write the <font> element for comments.
    fn write_comment_font(&mut self) {
        xml_start_tag_only(&mut self.writer, "font");

        xml_empty_tag(&mut self.writer, "sz", &[("val", "8")]);
        xml_empty_tag(&mut self.writer, "color", &[("indexed", "81")]);
        xml_empty_tag(&mut self.writer, "name", &[("val", "Tahoma")]);
        xml_empty_tag(&mut self.writer, "family", &[("val", "2")]);

        xml_end_tag(&mut self.writer, "font");
    }

    // Write the <vertAlign> element.
    fn write_vert_align(&mut self, font: &Font) {
        let mut attributes = vec![];

        match font.script {
            FormatScript::Superscript => {
                attributes.push(("val", "superscript".to_string()));
            }
            FormatScript::Subscript => {
                attributes.push(("val", "subscript".to_string()));
            }
            FormatScript::None => {}
        }

        xml_empty_tag(&mut self.writer, "vertAlign", &attributes);
    }

    // Write the <fills> element.
    fn write_fills(&mut self) {
        let attributes = [("count", self.fill_count.to_string())];

        xml_start_tag(&mut self.writer, "fills", &attributes);

        // Write the default fill elements.
        self.write_default_fill("none".to_string());
        self.write_default_fill("gray125".to_string());

        // Write the cell fill elements.
        for xf_format in self.xf_formats {
            // Write the fill element.
            if xf_format.has_fill {
                self.write_fill(&xf_format.fill, false);
            }
        }

        xml_end_tag(&mut self.writer, "fills");
    }

    // Write the default <fill> element.
    fn write_default_fill(&mut self, pattern: String) {
        let attributes = [("patternType", pattern)];

        xml_start_tag_only(&mut self.writer, "fill");
        xml_empty_tag(&mut self.writer, "patternFill", &attributes);
        xml_end_tag(&mut self.writer, "fill");
    }

    // Write the user-defined <fill> element.
    fn write_fill(&mut self, fill: &Fill, dxf_format: bool) {
        // Special handling for pattern only case.
        if fill.pattern != FormatPattern::None
            && (fill.background_color == Color::Default
                || fill.background_color == Color::Automatic)
            && (fill.foreground_color == Color::Default
                || fill.foreground_color == Color::Automatic)
        {
            self.write_default_fill(fill.pattern.to_string());
            return;
        }

        // Excel reverses the role of foreground and background colors for a
        // solid fill in an XF format.
        let mut foreground_color_attributes = fill.foreground_color.attributes();
        let mut background_color_attributes = fill.background_color.attributes();
        if !dxf_format
            && fill.pattern == FormatPattern::Solid
            && fill.background_color != Color::Default
            && fill.foreground_color != Color::Default
        {
            mem::swap(
                &mut foreground_color_attributes,
                &mut background_color_attributes,
            );
        }

        // Start the "fill" element.
        xml_start_tag_only(&mut self.writer, "fill");

        // Write the fill pattern.
        let mut attributes = vec![("patternType", fill.pattern.to_string())];

        if dxf_format
            && (fill.pattern == FormatPattern::None || fill.pattern == FormatPattern::Solid)
        {
            attributes = vec![];
        }

        xml_start_tag(&mut self.writer, "patternFill", &attributes);

        // Write the foreground color.
        if fill.foreground_color != Color::Default && fill.foreground_color != Color::Automatic {
            xml_empty_tag(&mut self.writer, "fgColor", &foreground_color_attributes);
        }

        // Write the background color.
        if fill.background_color == Color::Default || fill.background_color == Color::Automatic {
            let attributes = [("indexed", "64")];
            xml_empty_tag(&mut self.writer, "bgColor", &attributes);
        } else {
            xml_empty_tag(&mut self.writer, "bgColor", &background_color_attributes);
        }

        xml_end_tag(&mut self.writer, "patternFill");
        xml_end_tag(&mut self.writer, "fill");
    }

    // Write the <borders> element.
    fn write_borders(&mut self) {
        let attributes = [("count", self.border_count.to_string())];

        xml_start_tag(&mut self.writer, "borders", &attributes);

        // Write the cell border elements.
        for xf_format in self.xf_formats {
            // Write the border element.
            if xf_format.has_border {
                self.write_border(&xf_format.borders, false);
            }
        }

        xml_end_tag(&mut self.writer, "borders");
    }

    // Write the <border> element.
    fn write_border(&mut self, borders: &Border, dxf_format: bool) {
        match borders.diagonal_type {
            FormatDiagonalBorder::None => {
                xml_start_tag_only(&mut self.writer, "border");
            }
            FormatDiagonalBorder::BorderUp => {
                let attributes = [("diagonalUp", "1")];
                xml_start_tag(&mut self.writer, "border", &attributes);
            }
            FormatDiagonalBorder::BorderDown => {
                let attributes = [("diagonalDown", "1")];
                xml_start_tag(&mut self.writer, "border", &attributes);
            }
            FormatDiagonalBorder::BorderUpDown => {
                let attributes = [("diagonalUp", "1"), ("diagonalDown", "1")];
                xml_start_tag(&mut self.writer, "border", &attributes);
            }
        }

        // Write the four border elements.
        self.write_sub_border("left", borders.left_style, borders.left_color);
        self.write_sub_border("right", borders.right_style, borders.right_color);
        self.write_sub_border("top", borders.top_style, borders.top_color);
        self.write_sub_border("bottom", borders.bottom_style, borders.bottom_color);

        if !dxf_format {
            self.write_sub_border("diagonal", borders.diagonal_style, borders.diagonal_color);
        }

        if dxf_format {
            xml_empty_tag_only(&mut self.writer, "vertical");
            xml_empty_tag_only(&mut self.writer, "horizontal");
        }

        xml_end_tag(&mut self.writer, "border");
    }

    // Write the <border> sub-elements such as <right>, <top>, etc.
    fn write_sub_border(
        &mut self,
        border_type: &str,
        border_style: FormatBorder,
        border_color: Color,
    ) {
        if border_style == FormatBorder::None {
            xml_empty_tag_only(&mut self.writer, border_type);
            return;
        }

        let mut attributes = vec![("style", border_style.to_string())];
        xml_start_tag(&mut self.writer, border_type, &attributes);

        if border_color != Color::Default && border_color != Color::Automatic {
            attributes = border_color.attributes();
        } else {
            attributes = vec![("auto", "1".to_string())];
        }

        xml_empty_tag(&mut self.writer, "color", &attributes);

        xml_end_tag(&mut self.writer, border_type);
    }

    // Write the <cellStyleXfs> element.
    fn write_cell_style_xfs(&mut self) {
        let mut count = 1;
        if self.has_hyperlink_style {
            count = 2;
        }

        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "cellStyleXfs", &attributes);

        // Write the style xf elements.
        self.write_normal_style_xf();

        if self.has_hyperlink_style {
            self.write_hyperlink_style_xf();
        }

        xml_end_tag(&mut self.writer, "cellStyleXfs");
    }

    // Write the style <xf> element for the "Normal" style.
    fn write_normal_style_xf(&mut self) {
        let attributes = [
            ("numFmtId", "0"),
            ("fontId", "0"),
            ("fillId", "0"),
            ("borderId", "0"),
        ];

        xml_empty_tag(&mut self.writer, "xf", &attributes);
    }

    // Write the style <xf> element for the "Hyperlink" style.
    fn write_hyperlink_style_xf(&mut self) {
        let attributes = [
            ("numFmtId", "0".to_string()),
            ("fontId", self.hyperlink_font_id.to_string()),
            ("fillId", "0".to_string()),
            ("borderId", "0".to_string()),
            ("applyNumberFormat", "0".to_string()),
            ("applyFill", "0".to_string()),
            ("applyBorder", "0".to_string()),
            ("applyAlignment", "0".to_string()),
            ("applyProtection", "0".to_string()),
        ];

        xml_start_tag(&mut self.writer, "xf", &attributes);
        self.write_hyperlink_alignment();
        self.write_hyperlink_protection();
        xml_end_tag(&mut self.writer, "xf");
    }

    // Write the <alignment> element for hyperlinks.
    fn write_hyperlink_alignment(&mut self) {
        let attributes = [("vertical", "top")];

        xml_empty_tag(&mut self.writer, "alignment", &attributes);
    }

    // Write the <protection> element for hyperlinks.
    fn write_hyperlink_protection(&mut self) {
        let attributes = [("locked", "0")];

        xml_empty_tag(&mut self.writer, "protection", &attributes);
    }

    // Write the <cellXfs> element.
    fn write_cell_xfs(&mut self) {
        let xf_count = format!("{}", self.xf_formats.len());
        let attributes = [("count", xf_count)];

        xml_start_tag(&mut self.writer, "cellXfs", &attributes);

        // Write the cell xf element.
        for xf_format in self.xf_formats {
            self.write_cell_xf(xf_format);
        }

        xml_end_tag(&mut self.writer, "cellXfs");
    }

    // Write the cell <xf> element.
    fn write_cell_xf(&mut self, xf_format: &Format) {
        let has_checkbox = xf_format.has_checkbox();
        let is_hyperlink = xf_format.font.is_hyperlink;
        let has_alignment = xf_format.has_alignment();
        let has_protection = xf_format.has_protection();
        let apply_alignment = xf_format.apply_alignment();
        let xf_id = i32::from(is_hyperlink);

        let mut attributes = vec![
            ("numFmtId", xf_format.num_format_index.to_string()),
            ("fontId", xf_format.font_index.to_string()),
            ("fillId", xf_format.fill_index.to_string()),
            ("borderId", xf_format.border_index.to_string()),
            ("xfId", xf_id.to_string()),
        ];

        if xf_format.quote_prefix {
            attributes.push(("quotePrefix", "1".to_string()));
        }

        if xf_format.num_format_index > 0 {
            attributes.push(("applyNumberFormat", "1".to_string()));
        }

        if xf_format.font_index > 0 && !is_hyperlink {
            attributes.push(("applyFont", "1".to_string()));
        }

        if xf_format.fill_index > 0 {
            attributes.push(("applyFill", "1".to_string()));
        }

        if xf_format.border_index > 0 {
            attributes.push(("applyBorder", "1".to_string()));
        }

        if apply_alignment || is_hyperlink {
            attributes.push(("applyAlignment", "1".to_string()));
        }

        if has_protection || is_hyperlink {
            attributes.push(("applyProtection", "1".to_string()));
        }

        if has_alignment || has_protection || has_checkbox {
            xml_start_tag(&mut self.writer, "xf", &attributes);

            if has_alignment {
                // Write the alignment element.
                self.write_alignment(xf_format.alignment);
            }

            if has_protection {
                // Write the protection element.
                self.write_protection(xf_format);
            }

            if has_checkbox {
                // Write the checkbox extLst element.
                self.write_xf_format_extensions();
            }

            xml_end_tag(&mut self.writer, "xf");
        } else {
            xml_empty_tag(&mut self.writer, "xf", &attributes);
        }
    }

    // Write the <protection> element.
    fn write_protection(&mut self, xf_format: &Format) {
        let mut attributes = vec![];

        if !xf_format.locked {
            attributes.push(("locked", "0".to_string()));
        }

        if xf_format.hidden {
            attributes.push(("hidden", "1".to_string()));
        }

        xml_empty_tag(&mut self.writer, "protection", &attributes);
    }

    // Write the <alignment> element.
    fn write_alignment(&mut self, alignment: Alignment) {
        let mut attributes = vec![];
        let mut horizontal_align = alignment.horizontal;
        let mut shrink = alignment.shrink;

        // Indent is only allowed for some alignment properties. If it is
        // defined for any other alignment or no alignment has been set then
        // default to left alignment.
        if alignment.indent > 0
            && horizontal_align != FormatAlign::Left
            && horizontal_align != FormatAlign::Right
            && horizontal_align != FormatAlign::Distributed
            && alignment.vertical != FormatAlign::Top
            && alignment.vertical != FormatAlign::Bottom
            && alignment.vertical != FormatAlign::Distributed
        {
            horizontal_align = FormatAlign::Left;
        }

        // Check for properties that are mutually exclusive with "shrink".
        if alignment.text_wrap
            || horizontal_align == FormatAlign::Fill
            || horizontal_align == FormatAlign::Justify
            || horizontal_align == FormatAlign::Distributed
        {
            shrink = false;
        }

        // Set the various attributes for horizontal alignment.
        match horizontal_align {
            FormatAlign::Center => {
                attributes.push(("horizontal", "center".to_string()));
            }
            FormatAlign::CenterAcross => {
                attributes.push(("horizontal", "centerContinuous".to_string()));
            }
            FormatAlign::Distributed => {
                attributes.push(("horizontal", "distributed".to_string()));
            }
            FormatAlign::Fill => {
                attributes.push(("horizontal", "fill".to_string()));
            }
            FormatAlign::Justify => {
                attributes.push(("horizontal", "justify".to_string()));
            }
            FormatAlign::Left => {
                attributes.push(("horizontal", "left".to_string()));
            }
            FormatAlign::Right => {
                attributes.push(("horizontal", "right".to_string()));
            }
            _ => {}
        }

        // Set the various attributes for vertical alignment.
        match alignment.vertical {
            FormatAlign::VerticalCenter => {
                attributes.push(("vertical", "center".to_string()));
            }
            FormatAlign::VerticalDistributed => {
                attributes.push(("vertical", "distributed".to_string()));
            }
            FormatAlign::VerticalJustify => {
                attributes.push(("vertical", "justify".to_string()));
            }
            FormatAlign::Top => {
                attributes.push(("vertical", "top".to_string()));
            }
            _ => {}
        }

        if alignment.rotation != 0 {
            attributes.push(("textRotation", alignment.rotation.to_string()));
        }

        // Set other alignment properties.
        if alignment.indent != 0 {
            attributes.push(("indent", alignment.indent.to_string()));
        }

        if alignment.text_wrap {
            attributes.push(("wrapText", "1".to_string()));
        }

        if shrink {
            attributes.push(("shrinkToFit", "1".to_string()));
        }

        if alignment.reading_direction > 0 && alignment.reading_direction <= 2 {
            attributes.push(("readingOrder", alignment.reading_direction.to_string()));
        }

        xml_empty_tag(&mut self.writer, "alignment", &attributes);
    }

    // Write the <cellStyles> element.
    fn write_cell_styles(&mut self) {
        let mut count = 1;
        if self.has_hyperlink_style {
            count = 2;
        }

        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "cellStyles", &attributes);

        // Write the cellStyle elements.
        if self.has_hyperlink_style {
            self.write_hyperlink_cell_style();
        }
        self.write_normal_cell_style();

        xml_end_tag(&mut self.writer, "cellStyles");
    }

    // Write the <cellStyle> element for the "Normal" style.
    fn write_normal_cell_style(&mut self) {
        let attributes = [("name", "Normal"), ("xfId", "0"), ("builtinId", "0")];

        xml_empty_tag(&mut self.writer, "cellStyle", &attributes);
    }

    // Write the <cellStyle> element for the "Hyperlink" style.
    fn write_hyperlink_cell_style(&mut self) {
        let attributes = [("name", "Hyperlink"), ("xfId", "1"), ("builtinId", "8")];

        xml_empty_tag(&mut self.writer, "cellStyle", &attributes);
    }

    // Write the <dxfs> element.
    fn write_dxfs(&mut self) {
        let attributes = [("count", self.dxf_formats.len().to_string())];

        if self.dxf_formats.is_empty() {
            xml_empty_tag(&mut self.writer, "dxfs", &attributes);
        } else {
            xml_start_tag(&mut self.writer, "dxfs", &attributes);

            for dxf_format in self.dxf_formats {
                xml_start_tag_only(&mut self.writer, "dxf");

                if dxf_format.has_dxf_font() {
                    self.write_font(&dxf_format.font, true);
                }

                if dxf_format.num_format_index > 0 {
                    self.write_num_fmt(dxf_format.num_format_index, &dxf_format.num_format);
                }

                if dxf_format.has_dxf_fill() {
                    self.write_fill(&dxf_format.fill, true);
                }

                if dxf_format.has_border {
                    self.write_border(&dxf_format.borders, true);
                }

                if dxf_format.has_checkbox() {
                    self.write_dxf_format_extensions();
                }

                xml_end_tag(&mut self.writer, "dxf");
            }

            xml_end_tag(&mut self.writer, "dxfs");
        }
    }

    // Write the <tableStyles> element.
    fn write_table_styles(&mut self) {
        let attributes = [
            ("count", "0"),
            ("defaultTableStyle", "TableStyleMedium9"),
            ("defaultPivotStyle", "PivotStyleLight16"),
        ];

        xml_empty_tag(&mut self.writer, "tableStyles", &attributes);
    }

    // Write the <numFmts> element.
    fn write_num_fmts(&mut self) {
        if self.num_formats.is_empty() {
            return;
        }

        let attributes = [("count", self.num_formats.len().to_string())];
        xml_start_tag(&mut self.writer, "numFmts", &attributes);

        // Write the numFmt elements.
        for (index, num_format) in self.num_formats.clone().iter().enumerate() {
            self.write_num_fmt(index as u16 + 164, num_format);
        }

        xml_end_tag(&mut self.writer, "numFmts");
    }

    // Write the <numFmt> element.
    fn write_num_fmt(&mut self, num_format_index: u16, num_format: &str) {
        let attributes = [
            ("numFmtId", num_format_index.to_string()),
            ("formatCode", num_format.to_string()),
        ];

        xml_empty_tag(&mut self.writer, "numFmt", &attributes);
    }

    // Write the xfComplement <extLst> elements.
    fn write_xf_format_extensions(&mut self) {
        let attributes = [
            ("uri", "{C7286773-470A-42A8-94C5-96B5CB345126}"),
            (
                "xmlns:xfpb",
                "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag",
            ),
        ];

        xml_start_tag_only(&mut self.writer, "extLst");
        xml_start_tag(&mut self.writer, "ext", &attributes);

        xml_empty_tag(&mut self.writer, "xfpb:xfComplement", &[("i", "0")]);

        xml_end_tag(&mut self.writer, "ext");
        xml_end_tag(&mut self.writer, "extLst");
    }

    // Write the DXFComplement <extLst> elements.
    fn write_dxf_format_extensions(&mut self) {
        let attributes = [
            ("uri", "{0417FA29-78FA-4A13-93AC-8FF0FAFDF519}"),
            (
                "xmlns:xfpb",
                "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag",
            ),
        ];

        xml_start_tag_only(&mut self.writer, "extLst");
        xml_start_tag(&mut self.writer, "ext", &attributes);

        xml_empty_tag(&mut self.writer, "xfpb:DXFComplement", &[("i", "0")]);

        xml_end_tag(&mut self.writer, "ext");
        xml_end_tag(&mut self.writer, "extLst");
    }
}
