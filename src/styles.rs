// styles - A module for creating the Excel styles.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::format::Format;
use crate::xmlwriter::XMLWriter;

pub struct Styles<'a> {
    pub writer: XMLWriter,
    xf_formats: &'a Vec<Format>,
    font_count: u32,
}

impl<'a> Styles<'a> {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Styles struct.
    pub(crate) fn new(xf_formats: &Vec<Format>, font_count: u32) -> Styles {
        let writer = XMLWriter::new();

        Styles {
            writer,
            xf_formats,
            font_count,
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

        // Write the cell xf element.
        for xf_format in self.xf_formats {
            // Write the font element.
            if xf_format.has_font() {
                self.write_font(xf_format);
            }
        }

        self.writer.xml_end_tag("fonts");
    }

    // Write the <font> element.
    fn write_font(&mut self, xf_format: &Format) {
        self.writer.xml_start_tag("font");

        if xf_format.bold() {
            self.writer.xml_empty_tag("b");
        }

        if xf_format.italic() {
            self.writer.xml_empty_tag("i");
        }

        // Write the sz element.
        self.write_sz();

        // Write the color element.
        self.write_color();

        // Write the name element.
        self.write_name();

        // Write the family element.
        self.write_family();

        // Write the scheme element.
        self.write_scheme();

        self.writer.xml_end_tag("font");
    }

    // Write the <sz> element.
    fn write_sz(&mut self) {
        let attributes = vec![("val", "11".to_string())];

        self.writer.xml_empty_tag_attr("sz", &attributes);
    }

    // Write the <color> element.
    fn write_color(&mut self) {
        let attributes = vec![("theme", "1".to_string())];

        self.writer.xml_empty_tag_attr("color", &attributes);
    }

    // Write the <name> element.
    fn write_name(&mut self) {
        let attributes = vec![("val", "Calibri".to_string())];

        self.writer.xml_empty_tag_attr("name", &attributes);
    }

    // Write the <family> element.
    fn write_family(&mut self) {
        let attributes = vec![("val", "2".to_string())];

        self.writer.xml_empty_tag_attr("family", &attributes);
    }

    // Write the <scheme> element.
    fn write_scheme(&mut self) {
        let attributes = vec![("val", "minor".to_string())];

        self.writer.xml_empty_tag_attr("scheme", &attributes);
    }

    // Write the <fills> element.
    fn write_fills(&mut self) {
        let attributes = vec![("count", "2".to_string())];

        self.writer.xml_start_tag_attr("fills", &attributes);

        // Write the default fill elements.
        self.write_default_fill("none".to_string());
        self.write_default_fill("gray125".to_string());

        self.writer.xml_end_tag("fills");
    }

    // Write the default <fill> element.
    fn write_default_fill(&mut self, pattern: String) {
        let attributes = vec![("patternType", pattern)];

        self.writer.xml_start_tag("fill");
        self.writer.xml_empty_tag_attr("patternFill", &attributes);
        self.writer.xml_end_tag("fill");
    }

    // Write the <borders> element.
    fn write_borders(&mut self) {
        let attributes = vec![("count", "1".to_string())];

        self.writer.xml_start_tag_attr("borders", &attributes);

        // Write the border element.
        self.write_border();

        self.writer.xml_end_tag("borders");
    }

    // Write the <border> element.
    fn write_border(&mut self) {
        self.writer.xml_start_tag("border");
        // Write the left element.
        self.write_left();

        // Write the right element.
        self.write_right();

        // Write the top element.
        self.write_top();

        // Write the bottom element.
        self.write_bottom();

        // Write the diagonal element.
        self.write_diagonal();

        self.writer.xml_end_tag("border");
    }

    // Write the <left> element.
    fn write_left(&mut self) {
        self.writer.xml_empty_tag("left");
    }

    // Write the <right> element.
    fn write_right(&mut self) {
        self.writer.xml_empty_tag("right");
    }

    // Write the <top> element.
    fn write_top(&mut self) {
        self.writer.xml_empty_tag("top");
    }

    // Write the <bottom> element.
    fn write_bottom(&mut self) {
        self.writer.xml_empty_tag("bottom");
    }

    // Write the <diagonal> element.
    fn write_diagonal(&mut self) {
        self.writer.xml_empty_tag("diagonal");
    }

    // Write the <cellStyleXfs> element.
    fn write_cell_style_xfs(&mut self) {
        let attributes = vec![("count", "1".to_string())];

        self.writer.xml_start_tag_attr("cellStyleXfs", &attributes);

        // Write the style xf element.
        self.write_style_xf();

        self.writer.xml_end_tag("cellStyleXfs");
    }

    // Write the style <xf> element.
    fn write_style_xf(&mut self) {
        let attributes = vec![
            ("numFmtId", "0".to_string()),
            ("fontId", "0".to_string()),
            ("fillId", "0".to_string()),
            ("borderId", "0".to_string()),
        ];

        self.writer.xml_empty_tag_attr("xf", &attributes);
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
        let mut attributes = vec![
            ("numFmtId", "0".to_string()),
            ("fontId", xf_format.get_font_index().to_string()),
            ("fillId", "0".to_string()),
            ("borderId", "0".to_string()),
            ("xfId", "0".to_string()),
        ];

        if xf_format.get_font_index() > 0 {
            attributes.push(("applyFont", "1".to_string()));
        }

        self.writer.xml_empty_tag_attr("xf", &attributes);
    }

    // Write the <cellStyles> element.
    fn write_cell_styles(&mut self) {
        let attributes = vec![("count", "1".to_string())];

        self.writer.xml_start_tag_attr("cellStyles", &attributes);

        // Write the cellStyle element.
        self.write_cell_style();

        self.writer.xml_end_tag("cellStyles");
    }

    // Write the <cellStyle> element.
    fn write_cell_style(&mut self) {
        let attributes = vec![
            ("name", "Normal".to_string()),
            ("xfId", "0".to_string()),
            ("builtinId", "0".to_string()),
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
}

#[cfg(test)]
mod tests {

    use super::Format;
    use super::Styles;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut xf_format = Format::new();
        xf_format.set_font_index(0, true);

        let xf_formats = vec![xf_format];
        let mut styles = Styles::new(&xf_formats, 1);

        styles.assemble_xml_file();

        let got = styles.writer.read_to_string();
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

        assert_eq!(got, expected);
    }
}
