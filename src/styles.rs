// styles - A module for creating the Excel styles.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::format::Format;
use crate::xmlwriter::XMLWriter;
use crate::{
    FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern, FormatScript, FormatUnderline,
    XlsxColor,
};

pub struct Styles<'a> {
    pub(crate) writer: XMLWriter,
    xf_formats: &'a Vec<Format>,
    font_count: u16,
    fill_count: u16,
    border_count: u16,
    num_formats: Vec<String>,
    has_hyperlink_style: bool,
    is_rich_string_style: bool,
}

impl<'a> Styles<'a> {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Styles struct.
    pub(crate) fn new(
        xf_formats: &Vec<Format>,
        font_count: u16,
        fill_count: u16,
        border_count: u16,
        num_formats: Vec<String>,
        has_hyperlink_style: bool,
        is_rich_string_style: bool,
    ) -> Styles {
        let writer = XMLWriter::new();

        Styles {
            writer,
            xf_formats,
            font_count,
            fill_count,
            border_count,
            num_formats,
            has_hyperlink_style,
            is_rich_string_style,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

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
        self.writer.xml_end_tag("styleSheet");
    }

    // Write the <styleSheet> element.
    fn write_style_sheet(&mut self) {
        let attributes = vec![(
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
        )];

        self.writer.xml_start_tag_attr("styleSheet", &attributes);
    }

    // Write the <fonts> element.
    fn write_fonts(&mut self) {
        let attributes = vec![("count", self.font_count.to_string())];

        self.writer.xml_start_tag_attr("fonts", &attributes);

        // Write the cell font elements.
        for xf_format in self.xf_formats {
            // Write the font element.
            if xf_format.has_font {
                self.write_font(xf_format);
            }
        }

        self.writer.xml_end_tag("fonts");
    }

    // Write the <font> element.
    pub(crate) fn write_font(&mut self, xf_format: &Format) {
        if self.is_rich_string_style {
            self.writer.xml_start_tag("rPr");
        } else {
            self.writer.xml_start_tag("font");
        }

        if xf_format.bold {
            self.writer.xml_empty_tag("b");
        }

        if xf_format.italic {
            self.writer.xml_empty_tag("i");
        }

        if xf_format.font_strikethrough {
            self.writer.xml_empty_tag("strike");
        }

        if xf_format.underline != FormatUnderline::None {
            self.write_font_underline(xf_format);
        }

        if xf_format.font_script != FormatScript::None {
            self.write_vert_align(xf_format);
        }
        // Write the sz element.
        self.write_font_size(xf_format);

        // Write the color element.
        self.write_font_color(xf_format);

        // Write the name element.
        self.write_font_name(xf_format);

        // Write the family element.
        if xf_format.font_family > 0 {
            self.write_font_family(xf_format);
        }

        // Write the charset element.
        if xf_format.font_charset > 0 {
            self.write_font_charset(xf_format);
        }

        // Write the scheme element.
        self.write_font_scheme(xf_format);

        if self.is_rich_string_style {
            self.writer.xml_end_tag("rPr");
        } else {
            self.writer.xml_end_tag("font");
        }
    }

    // Write the <sz> element.
    fn write_font_size(&mut self, xf_format: &Format) {
        let attributes = vec![("val", xf_format.font_size.to_string())];

        self.writer.xml_empty_tag_attr("sz", &attributes);
    }

    // Write the <color> element.
    fn write_font_color(&mut self, xf_format: &Format) {
        let mut attributes = vec![];

        match xf_format.font_color {
            XlsxColor::Automatic => {
                attributes.push(("theme", "1".to_string()));
            }
            _ => {
                attributes.append(&mut xf_format.font_color.attributes());
            }
        }

        self.writer.xml_empty_tag_attr("color", &attributes);
    }

    // Write the <name> element.
    fn write_font_name(&mut self, xf_format: &Format) {
        let attributes = vec![("val", xf_format.font_name.clone())];

        if self.is_rich_string_style {
            self.writer.xml_empty_tag_attr("rFont", &attributes);
        } else {
            self.writer.xml_empty_tag_attr("name", &attributes);
        }
    }

    // Write the <family> element.
    fn write_font_family(&mut self, xf_format: &Format) {
        let attributes = vec![("val", xf_format.font_family.to_string())];

        self.writer.xml_empty_tag_attr("family", &attributes);
    }

    // Write the <charset> element.
    fn write_font_charset(&mut self, xf_format: &Format) {
        let attributes = vec![("val", xf_format.font_charset.to_string())];

        self.writer.xml_empty_tag_attr("charset", &attributes);
    }

    // Write the <scheme> element.
    fn write_font_scheme(&mut self, xf_format: &Format) {
        let mut attributes = vec![];

        if !xf_format.font_scheme.is_empty() {
            attributes.push(("val", xf_format.font_scheme.to_string()));
        } else {
            return;
        }

        self.writer.xml_empty_tag_attr("scheme", &attributes);
    }

    // Write the <u> underline element.
    fn write_font_underline(&mut self, xf_format: &Format) {
        let mut attributes = vec![];

        match xf_format.underline {
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

        self.writer.xml_empty_tag_attr("u", &attributes);
    }

    // Write the <vertAlign> element.
    fn write_vert_align(&mut self, xf_format: &Format) {
        let mut attributes = vec![];

        match xf_format.font_script {
            FormatScript::Superscript => {
                attributes.push(("val", "superscript".to_string()));
            }
            FormatScript::Subscript => {
                attributes.push(("val", "subscript".to_string()));
            }
            _ => {}
        }

        self.writer.xml_empty_tag_attr("vertAlign", &attributes);
    }

    // Write the <fills> element.
    fn write_fills(&mut self) {
        let attributes = vec![("count", self.fill_count.to_string())];

        self.writer.xml_start_tag_attr("fills", &attributes);

        // Write the default fill elements.
        self.write_default_fill("none".to_string());
        self.write_default_fill("gray125".to_string());

        // Write the cell fill elements.
        for xf_format in self.xf_formats {
            // Write the fill element.
            if xf_format.has_fill {
                self.write_fill(xf_format);
            }
        }

        self.writer.xml_end_tag("fills");
    }

    // Write the default <fill> element.
    fn write_default_fill(&mut self, pattern: String) {
        let attributes = vec![("patternType", pattern)];

        self.writer.xml_start_tag("fill");
        self.writer.xml_empty_tag_attr("patternFill", &attributes);
        self.writer.xml_end_tag("fill");
    }

    // Write the user defined <fill> element.
    fn write_fill(&mut self, xf_format: &Format) {
        // Special handling for pattern only case.
        if xf_format.pattern != FormatPattern::None
            && xf_format.background_color.is_default()
            && xf_format.foreground_color.is_default()
        {
            self.write_default_fill(xf_format.pattern.value().to_string());
            return;
        }

        // Start the "fill" element.
        self.writer.xml_start_tag("fill");

        // Write the fill pattern.
        let attributes = vec![("patternType", xf_format.pattern.value().to_string())];
        self.writer.xml_start_tag_attr("patternFill", &attributes);

        // Write the foreground color.
        if xf_format.foreground_color.is_not_default() {
            let attributes = xf_format.foreground_color.attributes();
            self.writer.xml_empty_tag_attr("fgColor", &attributes);
        }

        // Write the background color.
        if xf_format.background_color.is_not_default() {
            let attributes = xf_format.background_color.attributes();
            self.writer.xml_empty_tag_attr("bgColor", &attributes);
        } else {
            let attributes = vec![("indexed", "64".to_string())];
            self.writer.xml_empty_tag_attr("bgColor", &attributes);
        }

        self.writer.xml_end_tag("patternFill");
        self.writer.xml_end_tag("fill");
    }

    // Write the <borders> element.
    fn write_borders(&mut self) {
        let attributes = vec![("count", self.border_count.to_string())];

        self.writer.xml_start_tag_attr("borders", &attributes);

        // Write the cell border elements.
        for xf_format in self.xf_formats {
            // Write the border element.
            if xf_format.has_border {
                self.write_border(xf_format);
            }
        }

        self.writer.xml_end_tag("borders");
    }

    // Write the <border> element.
    fn write_border(&mut self, xf_format: &Format) {
        match xf_format.border_diagonal_type {
            FormatDiagonalBorder::None => {
                self.writer.xml_start_tag("border");
            }
            FormatDiagonalBorder::BorderUp => {
                let attributes = vec![("diagonalUp", "1".to_string())];
                self.writer.xml_start_tag_attr("border", &attributes);
            }
            FormatDiagonalBorder::BorderDown => {
                let attributes = vec![("diagonalDown", "1".to_string())];
                self.writer.xml_start_tag_attr("border", &attributes);
            }
            FormatDiagonalBorder::BorderUpDown => {
                let attributes = vec![
                    ("diagonalUp", "1".to_string()),
                    ("diagonalDown", "1".to_string()),
                ];
                self.writer.xml_start_tag_attr("border", &attributes);
            }
        }

        // Write the four border elements.
        self.write_sub_border("left", xf_format.border_left, xf_format.border_left_color);
        self.write_sub_border(
            "right",
            xf_format.border_right,
            xf_format.border_right_color,
        );
        self.write_sub_border("top", xf_format.border_top, xf_format.border_top_color);
        self.write_sub_border(
            "bottom",
            xf_format.border_bottom,
            xf_format.border_bottom_color,
        );
        self.write_sub_border(
            "diagonal",
            xf_format.border_diagonal,
            xf_format.border_diagonal_color,
        );

        self.writer.xml_end_tag("border");
    }

    // Write the <border> sub elements such as <right>, <top>, etc.
    fn write_sub_border(
        &mut self,
        border_type: &str,
        border_style: FormatBorder,
        border_color: XlsxColor,
    ) {
        if border_style == FormatBorder::None {
            self.writer.xml_empty_tag(border_type);
            return;
        }

        let mut attributes = vec![("style", border_style.value().to_string())];
        self.writer.xml_start_tag_attr(border_type, &attributes);

        if border_color.is_not_default() {
            attributes = border_color.attributes();
        } else {
            attributes = vec![("auto", "1".to_string())];
        }

        self.writer.xml_empty_tag_attr("color", &attributes);

        self.writer.xml_end_tag(border_type);
    }

    // Write the <cellStyleXfs> element.
    fn write_cell_style_xfs(&mut self) {
        let mut count = 1;
        if self.has_hyperlink_style {
            count = 2;
        }

        let attributes = vec![("count", count.to_string())];

        self.writer.xml_start_tag_attr("cellStyleXfs", &attributes);

        // Write the style xf elements.
        self.write_normal_style_xf();

        if self.has_hyperlink_style {
            self.write_hyperlink_style_xf();
        }

        self.writer.xml_end_tag("cellStyleXfs");
    }

    // Write the style <xf> element for the "Normal" style.
    fn write_normal_style_xf(&mut self) {
        let attributes = vec![
            ("numFmtId", "0".to_string()),
            ("fontId", "0".to_string()),
            ("fillId", "0".to_string()),
            ("borderId", "0".to_string()),
        ];

        self.writer.xml_empty_tag_attr("xf", &attributes);
    }

    // Write the style <xf> element for the "Hyperlink" style.
    fn write_hyperlink_style_xf(&mut self) {
        let attributes = vec![
            ("numFmtId", "0".to_string()),
            ("fontId", "1".to_string()),
            ("fillId", "0".to_string()),
            ("borderId", "0".to_string()),
            ("applyNumberFormat", "0".to_string()),
            ("applyFill", "0".to_string()),
            ("applyBorder", "0".to_string()),
            ("applyAlignment", "0".to_string()),
            ("applyProtection", "0".to_string()),
        ];

        self.writer.xml_start_tag_attr("xf", &attributes);
        self.write_hyperlink_alignment();
        self.write_hyperlink_protection();
        self.writer.xml_end_tag("xf");
    }

    // Write the <alignment> element for hyperlinks.
    fn write_hyperlink_alignment(&mut self) {
        let attributes = vec![("vertical", "top".to_string())];

        self.writer.xml_empty_tag_attr("alignment", &attributes);
    }

    // Write the <protection> element for hyperlinks.
    fn write_hyperlink_protection(&mut self) {
        let attributes = vec![("locked", "0".to_string())];

        self.writer.xml_empty_tag_attr("protection", &attributes);
    }

    // Write the <cellXfs> element.
    fn write_cell_xfs(&mut self) {
        let xf_count = format!("{}", self.xf_formats.len());
        let attributes = vec![("count", xf_count)];

        self.writer.xml_start_tag_attr("cellXfs", &attributes);

        // Write the cell xf element.
        for xf_format in self.xf_formats {
            self.write_cell_xf(xf_format);
        }

        self.writer.xml_end_tag("cellXfs");
    }

    // Write the cell <xf> element.
    fn write_cell_xf(&mut self, xf_format: &Format) {
        let has_protection = xf_format.has_protection();
        let has_alignment = xf_format.has_alignment();
        let apply_alignment = xf_format.apply_alignment();
        let is_hyperlink = xf_format.is_hyperlink;
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

        if has_alignment || has_protection {
            self.writer.xml_start_tag_attr("xf", &attributes);

            if has_alignment {
                // Write the alignment element.
                self.write_alignment(xf_format);
            }

            if has_protection {
                // Write the protection element.
                self.write_protection(xf_format);
            }

            self.writer.xml_end_tag("xf");
        } else {
            self.writer.xml_empty_tag_attr("xf", &attributes);
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

        self.writer.xml_empty_tag_attr("protection", &attributes);
    }

    // Write the <alignment> element.
    fn write_alignment(&mut self, xf_format: &Format) {
        let mut attributes = vec![];
        let mut horizontal_align = xf_format.horizontal_align;
        let mut shrink = xf_format.shrink;

        // Indent is only allowed for horizontal "left", "right" and
        // "distributed". If it is defined for any other alignment or no
        // alignment has been set then default to left alignment.
        if xf_format.indent > 0
            && horizontal_align != FormatAlign::Left
            && horizontal_align != FormatAlign::Right
            && horizontal_align != FormatAlign::Distributed
        {
            horizontal_align = FormatAlign::Left;
        }

        // Check for properties that are mutually exclusive with "shrink".
        if xf_format.text_wrap
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
        match xf_format.vertical_align {
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

        // Set other alignment properties.
        if xf_format.indent != 0 {
            attributes.push(("indent", xf_format.indent.to_string()));
        }

        if xf_format.rotation != 0 {
            attributes.push(("textRotation", xf_format.rotation.to_string()));
        }

        if xf_format.text_wrap {
            attributes.push(("wrapText", "1".to_string()));
        }

        if shrink {
            attributes.push(("shrinkToFit", "1".to_string()));
        }

        if xf_format.reading_direction > 0 && xf_format.reading_direction <= 2 {
            attributes.push(("readingOrder", xf_format.reading_direction.to_string()));
        }

        self.writer.xml_empty_tag_attr("alignment", &attributes);
    }

    // Write the <cellStyles> element.
    fn write_cell_styles(&mut self) {
        let mut count = 1;
        if self.has_hyperlink_style {
            count = 2;
        }

        let attributes = vec![("count", count.to_string())];

        self.writer.xml_start_tag_attr("cellStyles", &attributes);

        // Write the cellStyle elements.
        if self.has_hyperlink_style {
            self.write_hyperlink_cell_style();
        }
        self.write_normal_cell_style();

        self.writer.xml_end_tag("cellStyles");
    }

    // Write the <cellStyle> element for the "Normal" style.
    fn write_normal_cell_style(&mut self) {
        let attributes = vec![
            ("name", "Normal".to_string()),
            ("xfId", "0".to_string()),
            ("builtinId", "0".to_string()),
        ];

        self.writer.xml_empty_tag_attr("cellStyle", &attributes);
    }

    // Write the <cellStyle> element for the "Hyperlink" style.
    fn write_hyperlink_cell_style(&mut self) {
        let attributes = vec![
            ("name", "Hyperlink".to_string()),
            ("xfId", "1".to_string()),
            ("builtinId", "8".to_string()),
        ];

        self.writer.xml_empty_tag_attr("cellStyle", &attributes);
    }

    // Write the <dxfs> element.
    fn write_dxfs(&mut self) {
        let attributes = vec![("count", "0".to_string())];

        self.writer.xml_empty_tag_attr("dxfs", &attributes);
    }

    // Write the <tableStyles> element.
    fn write_table_styles(&mut self) {
        let attributes = vec![
            ("count", "0".to_string()),
            ("defaultTableStyle", "TableStyleMedium9".to_string()),
            ("defaultPivotStyle", "PivotStyleLight16".to_string()),
        ];

        self.writer.xml_empty_tag_attr("tableStyles", &attributes);
    }

    // Write the <numFmts> element.
    fn write_num_fmts(&mut self) {
        if self.num_formats.is_empty() {
            return;
        }

        let attributes = vec![("count", self.num_formats.len().to_string())];
        self.writer.xml_start_tag_attr("numFmts", &attributes);

        // Write the numFmt elements.
        for (index, num_format) in self.num_formats.clone().iter().enumerate() {
            self.write_num_fmt(index as u16 + 164, num_format);
        }

        self.writer.xml_end_tag("numFmts");
    }

    // Write the <numFmt> element.
    fn write_num_fmt(&mut self, num_format_index: u16, num_format: &str) {
        let attributes = vec![
            ("numFmtId", num_format_index.to_string()),
            ("formatCode", num_format.to_string()),
        ];

        self.writer.xml_empty_tag_attr("numFmt", &attributes);
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::styles::Styles;
    use crate::test_functions::xml_to_vec;
    use crate::Format;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut xf_format = Format::new();
        xf_format.set_font_index(0, true);
        xf_format.set_border_index(0, true);

        let xf_formats = vec![xf_format];
        let mut styles = Styles::new(&xf_formats, 1, 2, 1, vec![], false, false);

        styles.assemble_xml_file();

        let got = styles.writer.read_to_str();
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <fonts count="1">
                    <font>
                    <sz val="11"/>
                    <color theme="1"/>
                    <name val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                    </font>
                </fonts>
                <fills count="2">
                    <fill>
                    <patternFill patternType="none"/>
                    </fill>
                    <fill>
                    <patternFill patternType="gray125"/>
                    </fill>
                </fills>
                <borders count="1">
                    <border>
                    <left/>
                    <right/>
                    <top/>
                    <bottom/>
                    <diagonal/>
                    </border>
                </borders>
                <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                </cellStyleXfs>
                <cellXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                </cellXfs>
                <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                </cellStyles>
                <dxfs count="0"/>
                <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
                </styleSheet>
                "#,
        );

        assert_eq!(expected, got);
    }
}
