// drawing - A module for creating the Excel Drawing.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use crate::{
    xmlwriter::XMLWriter, Color, ObjectMovement, Shape, ShapeFont, ShapeFormat, ShapeGradientFill,
    ShapeGradientFillType, ShapeGradientStop, ShapeLine, ShapeLineDashType, ShapePatternFill,
    ShapeTextDirection, ShapeTextHorizontalAlignment, Url,
};

pub struct Drawing {
    pub(crate) writer: XMLWriter,
    pub(crate) drawings: Vec<DrawingInfo>,
    pub(crate) shapes: Vec<Shape>,
    shape_id: usize,
}

impl Drawing {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Drawing struct.
    pub fn new() -> Drawing {
        let writer = XMLWriter::new();

        Drawing {
            writer,
            drawings: vec![],
            shapes: vec![],
            shape_id: 0,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the xdr:wsDr element.
        self.write_ws_dr();

        for (index, drawing) in self.drawings.clone().iter().enumerate() {
            // Write the xdr:twoCellAnchor element.
            self.write_two_cell_anchor((index + 1) as u32, drawing);
        }

        // Close the end tag.
        self.writer.xml_end_tag("xdr:wsDr");
    }

    // Write the <xdr:wsDr> element.
    fn write_ws_dr(&mut self) {
        let attributes = [
            (
                "xmlns:xdr",
                "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            ),
            (
                "xmlns:a",
                "http://schemas.openxmlformats.org/drawingml/2006/main",
            ),
        ];

        self.writer.xml_start_tag("xdr:wsDr", &attributes);
    }

    // Write the <xdr:twoCellAnchor> element.
    fn write_two_cell_anchor(&mut self, index: u32, drawing_info: &DrawingInfo) {
        let mut attributes = vec![];

        match drawing_info.object_movement {
            ObjectMovement::MoveButDontSizeWithCells => {
                attributes.push(("editAs", "oneCell".to_string()));
            }
            ObjectMovement::DontMoveOrSizeWithCells => {
                attributes.push(("editAs", "absolute".to_string()));
            }
            ObjectMovement::MoveAndSizeWithCells | ObjectMovement::MoveAndSizeWithCellsAfter => (),
        }

        self.writer.xml_start_tag("xdr:twoCellAnchor", &attributes);

        // Write the xdr:from and xdr:to elements
        self.write_from(&drawing_info.from);
        self.write_to(&drawing_info.to);

        match drawing_info.drawing_type {
            DrawingType::Image => self.write_pic(index, drawing_info),
            DrawingType::Chart => self.write_graphic_frame(index, drawing_info),
            DrawingType::Shape => {
                let shape = self.shapes[self.shape_id].clone();
                self.shape_id += 1;

                self.write_sp(index, drawing_info, &shape);
            }
            DrawingType::Vml => {}
        }

        self.writer.xml_empty_tag_only("xdr:clientData");
        self.writer.xml_end_tag("xdr:twoCellAnchor");
    }

    // Write the <xdr:from> element.
    fn write_from(&mut self, coords: &DrawingCoordinates) {
        self.writer.xml_start_tag_only("xdr:from");

        self.writer
            .xml_data_element_only("xdr:col", &coords.col.to_string());
        self.writer
            .xml_data_element_only("xdr:colOff", &coords.col_offset.to_string());
        self.writer
            .xml_data_element_only("xdr:row", &coords.row.to_string());
        self.writer
            .xml_data_element_only("xdr:rowOff", &coords.row_offset.to_string());

        self.writer.xml_end_tag("xdr:from");
    }

    // Write the <xdr:to> element.
    fn write_to(&mut self, coords: &DrawingCoordinates) {
        self.writer.xml_start_tag_only("xdr:to");

        self.writer
            .xml_data_element_only("xdr:col", &coords.col.to_string());
        self.writer
            .xml_data_element_only("xdr:colOff", &coords.col_offset.to_string());
        self.writer
            .xml_data_element_only("xdr:row", &coords.row.to_string());
        self.writer
            .xml_data_element_only("xdr:rowOff", &coords.row_offset.to_string());

        self.writer.xml_end_tag("xdr:to");
    }

    // Write the <xdr:pic> element.
    fn write_pic(&mut self, index: u32, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag_only("xdr:pic");

        // Write the xdr:nvPicPr element.
        self.write_nv_pic_pr(index, drawing_info);

        // Write the xdr:blipFill element.
        self.write_blip_fill(drawing_info.rel_id);

        // Write the xdr:spPr element.
        self.write_sp_pr(drawing_info);

        self.writer.xml_end_tag("xdr:pic");
    }

    // Write the <xdr:nvPicPr> element.
    fn write_nv_pic_pr(&mut self, index: u32, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag_only("xdr:nvPicPr");

        // Write the xdr:cNvPr element.
        self.write_c_nv_pr(index, drawing_info, "Picture");

        // Write the xdr:cNvPicPr element.
        self.writer.xml_start_tag_only("xdr:cNvPicPr");
        self.write_a_pic_locks();
        self.writer.xml_end_tag("xdr:cNvPicPr");

        self.writer.xml_end_tag("xdr:nvPicPr");
    }

    // Write the <xdr:cNvPr> element.
    fn write_c_nv_pr(&mut self, index: u32, drawing_info: &DrawingInfo, name: &str) {
        let id = index + 1;
        let mut name = format!("{name} {index}");

        if !name.starts_with("TextBox") && !drawing_info.name.is_empty() {
            name.clone_from(&drawing_info.name);
        }

        let mut attributes = vec![("id", id.to_string()), ("name", name)];

        if !drawing_info.description.is_empty() {
            attributes.push(("descr", drawing_info.description.clone()));
        }

        if drawing_info.decorative || drawing_info.url.is_some() {
            self.writer.xml_start_tag("xdr:cNvPr", &attributes);

            if let Some(hyperlink) = &drawing_info.url {
                // Write the a:hlinkClick element.
                self.write_hyperlink(hyperlink);
            }

            if drawing_info.decorative {
                self.write_decorative();
            }

            self.writer.xml_end_tag("xdr:cNvPr");
        } else {
            self.writer.xml_empty_tag("xdr:cNvPr", &attributes);
        }
    }

    // Write the decorative sub elements.
    fn write_decorative(&mut self) {
        self.writer.xml_start_tag_only("a:extLst");

        let attributes = [("uri", "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}")];
        self.writer.xml_start_tag("a:ext", &attributes);

        let attributes = [
            (
                "xmlns:a16",
                "http://schemas.microsoft.com/office/drawing/2014/main",
            ),
            ("id", "{00000000-0008-0000-0000-000002000000}"),
        ];
        self.writer.xml_empty_tag("a16:creationId", &attributes);

        self.writer.xml_end_tag("a:ext");

        let attributes = [("uri", "{C183D7F6-B498-43B3-948B-1728B52AA6E4}")];
        self.writer.xml_start_tag("a:ext", &attributes);

        let attributes = [
            (
                "xmlns:adec",
                "http://schemas.microsoft.com/office/drawing/2017/decorative",
            ),
            ("val", "1"),
        ];
        self.writer.xml_empty_tag("adec:decorative", &attributes);

        self.writer.xml_end_tag("a:ext");
        self.writer.xml_end_tag("a:extLst");
    }

    // Write the <a:hlinkClick> element.
    fn write_hyperlink(&mut self, hyperlink: &Url) {
        let mut attributes = vec![
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            ),
            ("r:id", format!("rId{}", hyperlink.rel_id)),
        ];

        if !hyperlink.tool_tip.is_empty() {
            attributes.push(("tooltip", hyperlink.tool_tip.clone()));
        }

        self.writer.xml_empty_tag("a:hlinkClick", &attributes);
    }

    // Write the <a:picLocks> element.
    fn write_a_pic_locks(&mut self) {
        let attributes = [("noChangeAspect", "1")];

        self.writer.xml_empty_tag("a:picLocks", &attributes);
    }

    // Write the <xdr:blipFill> element.
    fn write_blip_fill(&mut self, index: u32) {
        self.writer.xml_start_tag_only("xdr:blipFill");

        // Write the a:blip element.
        self.write_a_blip(index);

        self.writer.xml_start_tag_only("a:stretch");
        self.writer.xml_empty_tag_only("a:fillRect");
        self.writer.xml_end_tag("a:stretch");

        self.writer.xml_end_tag("xdr:blipFill");
    }

    // Write the <a:blip> element.
    fn write_a_blip(&mut self, index: u32) {
        let attributes = [
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            ),
            ("r:embed", format!("rId{index}")),
        ];

        self.writer.xml_empty_tag("a:blip", &attributes);
    }

    // Write the <xdr:spPr> element.
    fn write_sp_pr(&mut self, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag_only("xdr:spPr");
        self.writer.xml_start_tag_only("a:xfrm");

        // Write the a:off element.
        self.write_a_off(drawing_info);

        // Write the a:ext element.
        self.write_a_ext(drawing_info);

        self.writer.xml_end_tag("a:xfrm");

        // Write the a:prstGeom element.
        self.write_a_prst_geom();

        self.writer.xml_end_tag("xdr:spPr");
    }

    // Write the <xdr:spPr> element.
    fn write_shape_sp_pr(&mut self, drawing_info: &DrawingInfo, shape: &Shape) {
        self.writer.xml_start_tag_only("xdr:spPr");
        self.writer.xml_start_tag_only("a:xfrm");

        // Write the a:off element.
        self.write_a_off(drawing_info);

        // Write the a:ext element.
        self.write_a_ext(drawing_info);

        self.writer.xml_end_tag("a:xfrm");

        // Write the a:prstGeom element.
        self.write_a_prst_geom();

        // Write the a:solidFill element.
        self.write_shape_formatting(&shape.format);

        self.writer.xml_end_tag("xdr:spPr");
    }

    // Write the <a:off> element.
    fn write_a_off(&mut self, drawing_info: &DrawingInfo) {
        let attributes = [
            ("x", drawing_info.col_absolute.to_string()),
            ("y", drawing_info.row_absolute.to_string()),
        ];

        self.writer.xml_empty_tag("a:off", &attributes);
    }

    // Write the <a:ext> element.
    fn write_a_ext(&mut self, drawing_info: &DrawingInfo) {
        let attributes = [
            ("cx", drawing_info.width.to_string()),
            ("cy", drawing_info.height.to_string()),
        ];

        self.writer.xml_empty_tag("a:ext", &attributes);
    }

    // Write the <a:prstGeom> element.
    fn write_a_prst_geom(&mut self) {
        let attributes = [("prst", "rect")];

        self.writer.xml_start_tag("a:prstGeom", &attributes);
        self.writer.xml_empty_tag_only("a:avLst");
        self.writer.xml_end_tag("a:prstGeom");
    }

    // Write the <xdr:graphicFrame> element.
    fn write_graphic_frame(&mut self, index: u32, drawing_info: &DrawingInfo) {
        let attributes = [("macro", "")];

        self.writer.xml_start_tag("xdr:graphicFrame", &attributes);

        // Write the xdr:nvGraphicFramePr element.
        self.write_nv_graphic_frame_pr(index, drawing_info);

        // Write the xdr:xfrm element.
        self.write_xfrm();

        // Write the a:graphic element.
        self.write_a_graphic(drawing_info.rel_id);

        self.writer.xml_end_tag("xdr:graphicFrame");
    }

    // Write the <xdr:nvGraphicFramePr> element.
    fn write_nv_graphic_frame_pr(&mut self, index: u32, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag_only("xdr:nvGraphicFramePr");

        // Write the xdr:cNvPr element.
        self.write_c_nv_pr(index, drawing_info, "Chart");

        // Write the xdr:cNvGraphicFramePr element.
        self.write_c_nv_graphic_frame_pr();

        self.writer.xml_end_tag("xdr:nvGraphicFramePr");
    }

    // Write the <xdr:cNvGraphicFramePr> element.
    fn write_c_nv_graphic_frame_pr(&mut self) {
        self.writer.xml_empty_tag_only("xdr:cNvGraphicFramePr");
    }

    // Write the <xdr:xfrm> element.
    fn write_xfrm(&mut self) {
        self.writer.xml_start_tag_only("xdr:xfrm");

        // Write the a:off element.
        self.write_chart_a_off();

        // Write the a:ext element.
        self.write_chart_a_ext();

        self.writer.xml_end_tag("xdr:xfrm");
    }

    // Write the <a:off> element.
    fn write_chart_a_off(&mut self) {
        let attributes = [("x", "0"), ("y", "0")];

        self.writer.xml_empty_tag("a:off", &attributes);
    }

    // Write the <a:ext> element.
    fn write_chart_a_ext(&mut self) {
        let attributes = [("cx", "0"), ("cy", "0")];

        self.writer.xml_empty_tag("a:ext", &attributes);
    }

    // Write the <a:graphic> element.
    fn write_a_graphic(&mut self, index: u32) {
        self.writer.xml_start_tag_only("a:graphic");

        // Write the a:graphicData element.
        self.write_a_graphic_data(index);

        self.writer.xml_end_tag("a:graphic");
    }

    // Write the <a:graphicData> element.
    fn write_a_graphic_data(&mut self, index: u32) {
        let attributes = [(
            "uri",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
        )];

        self.writer.xml_start_tag("a:graphicData", &attributes);

        // Write the c:chart element.
        self.write_chart(index);

        self.writer.xml_end_tag("a:graphicData");
    }

    // Write the <c:chart> element.
    fn write_chart(&mut self, index: u32) {
        let attributes = [
            (
                "xmlns:c",
                "http://schemas.openxmlformats.org/drawingml/2006/chart".to_string(),
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            ),
            ("r:id", format!("rId{index}")),
        ];

        self.writer.xml_empty_tag("c:chart", &attributes);
    }

    // Write the <xdr:sp> element.
    fn write_sp(&mut self, index: u32, drawing_info: &DrawingInfo, shape: &Shape) {
        let mut attributes = vec![("macro", String::new())];

        match &shape.text_link {
            Some(text_link) => {
                attributes.push(("textlink", text_link.formula_string.clone()));
            }
            None => {
                attributes.push(("textlink", String::new()));
            }
        }

        self.writer.xml_start_tag("xdr:sp", &attributes);

        // Write the xdr:nvSpPr element.
        self.write_nv_sp_pr(index, drawing_info);

        // Write the xdr:spPr element.
        self.write_shape_sp_pr(drawing_info, shape);

        // Write the xdr:style element.
        self.write_style();

        // Write the xdr:txBody element.
        self.write_tx_body(drawing_info, shape);

        self.writer.xml_end_tag("xdr:sp");
    }

    // Write the <xdr:nvSpPr> element.
    fn write_nv_sp_pr(&mut self, index: u32, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag_only("xdr:nvSpPr");

        // Write the xdr:cNvPr element.
        self.write_c_nv_pr(index, drawing_info, "TextBox");

        // Write the xdr:cNvSpPr element.
        self.write_c_nv_sp_pr();

        self.writer.xml_end_tag("xdr:nvSpPr");
    }

    // Write the <xdr:cNvSpPr> element.
    fn write_c_nv_sp_pr(&mut self) {
        let attributes = [("txBox", "1")];

        self.writer.xml_empty_tag("xdr:cNvSpPr", &attributes);
    }

    // Write the formatting elements for shapes.
    fn write_shape_formatting(&mut self, format: &ShapeFormat) {
        if format.no_fill {
            self.writer.xml_empty_tag_only("a:noFill");
        } else if let Some(solid_fill) = &format.solid_fill {
            // Write the a:solidFill element.
            self.write_a_solid_fill(solid_fill.color, solid_fill.transparency);
        } else if let Some(pattern_fill) = &format.pattern_fill {
            // Write the a:pattFill element.
            self.write_a_patt_fill(pattern_fill);
        } else if let Some(gradient_fill) = &format.gradient_fill {
            // Write the a:gradFill element.
            self.write_gradient_fill(gradient_fill);
        } else {
            // Write the a:solidFill element.
            self.write_default_solid_fill();
        }

        if format.no_line {
            // Write a default line with no fill.
            self.write_a_ln_none();
        } else if let Some(line) = &format.line {
            // Write the a:ln element.
            self.write_a_ln(line);
        } else {
            // Write the default a:ln element.
            let line = ShapeLine::new();
            self.write_a_ln(&line);
        }
    }

    // Write the <a:ln> element.
    fn write_a_ln(&mut self, line: &ShapeLine) {
        let mut attributes = vec![];

        // Round width to nearest 0.25, like Excel.
        let width = ((line.width + 0.125) * 4.0).floor() / 4.0;

        // Convert to Excel internal units.
        let width = (12700.0 * width).ceil() as u32;

        attributes.push(("w", width.to_string()));
        attributes.push(("cmpd", "sng".to_string()));

        self.writer.xml_start_tag("a:ln", &attributes);

        if line.color != Color::Default || line.dash_type != ShapeLineDashType::Solid || line.hidden
        {
            if line.hidden {
                // Write the a:noFill element.
                self.write_a_no_fill();
            } else {
                if line.color == Color::Default {
                    // Write the a:solidFill element.
                    self.write_line_solid_fill();
                } else {
                    // Write the a:solidFill element.
                    self.write_a_solid_fill(line.color, line.transparency);
                }

                if line.dash_type != ShapeLineDashType::Solid {
                    // Write the a:prstDash element.
                    self.write_a_prst_dash(line);
                }
            }
        } else {
            // Write the a:solidFill element.
            self.write_line_solid_fill();
        }

        self.writer.xml_end_tag("a:ln");
    }

    // Write the <a:ln> element.
    fn write_a_ln_none(&mut self) {
        let attributes = [("w", "9525"), ("cmpd", "sng")];

        self.writer.xml_start_tag("a:ln", &attributes);

        // Write the a:noFill element.
        self.write_a_no_fill();

        self.writer.xml_end_tag("a:ln");
    }

    // Write the <a:solidFill> element for the ln element.
    fn write_line_solid_fill(&mut self) {
        self.writer.xml_start_tag_only("a:solidFill");

        // Write the a:schemeClr element.
        self.write_default_scheme_clr("lt1", true);

        self.writer.xml_end_tag("a:solidFill");
    }

    // Write the default <a:solidFill> element.
    fn write_default_solid_fill(&mut self) {
        self.writer.xml_start_tag_only("a:solidFill");

        self.write_default_scheme_clr("lt1", false);

        self.writer.xml_end_tag("a:solidFill");
    }

    // Write the <a:solidFill> element.
    fn write_a_solid_fill(&mut self, color: Color, transparency: u8) {
        self.writer.xml_start_tag_only("a:solidFill");

        // Write the color element.
        self.write_color(color, transparency);

        self.writer.xml_end_tag("a:solidFill");
    }

    // Write the <a:pattFill> element.
    fn write_a_patt_fill(&mut self, fill: &ShapePatternFill) {
        let attributes = [("prst", fill.pattern.to_string())];

        self.writer.xml_start_tag("a:pattFill", &attributes);

        if fill.foreground_color != Color::Default {
            // Write the <a:fgClr> element.
            self.writer.xml_start_tag_only("a:fgClr");
            self.write_color(fill.foreground_color, 0);
            self.writer.xml_end_tag("a:fgClr");
        }

        if fill.background_color != Color::Default {
            // Write the <a:bgClr> element.
            self.writer.xml_start_tag_only("a:bgClr");
            self.write_color(fill.background_color, 0);
            self.writer.xml_end_tag("a:bgClr");
        } else if fill.background_color == Color::Default && fill.foreground_color != Color::Default
        {
            // If there is a foreground color but no background color then we
            // need to write a default background color.
            self.writer.xml_start_tag_only("a:bgClr");
            self.write_color(Color::White, 0);
            self.writer.xml_end_tag("a:bgClr");
        }

        self.writer.xml_end_tag("a:pattFill");
    }

    // Write the <a:gradFill> element.
    fn write_gradient_fill(&mut self, fill: &ShapeGradientFill) {
        let mut attributes = vec![];

        if fill.gradient_type != ShapeGradientFillType::Linear {
            attributes.push(("flip", "none"));
            attributes.push(("rotWithShape", "1"));
        }

        self.writer.xml_start_tag("a:gradFill", &attributes);
        self.writer.xml_start_tag_only("a:gsLst");

        for gradient_stop in &fill.gradient_stops {
            // Write the a:gs element.
            self.write_gradient_stop(gradient_stop);
        }

        self.writer.xml_end_tag("a:gsLst");

        if fill.gradient_type == ShapeGradientFillType::Linear {
            // Write the a:lin element.
            self.write_gradient_fill_angle(fill.angle);
        } else {
            // Write the a:path element.
            self.write_gradient_path(fill.gradient_type);
        }

        self.writer.xml_end_tag("a:gradFill");
    }

    // Write the <a:gs> element.
    fn write_gradient_stop(&mut self, gradient_stop: &ShapeGradientStop) {
        let position = 1000 * u32::from(gradient_stop.position);
        let attributes = [("pos", position.to_string())];

        self.writer.xml_start_tag("a:gs", &attributes);
        self.write_color(gradient_stop.color, 0);

        self.writer.xml_end_tag("a:gs");
    }

    // Write the <a:lin> element.
    fn write_gradient_fill_angle(&mut self, angle: u16) {
        let angle = 60_000 * u32::from(angle);
        let attributes = [("ang", angle.to_string()), ("scaled", "0".to_string())];

        self.writer.xml_empty_tag("a:lin", &attributes);
    }

    // Write the <a:path> element.
    fn write_gradient_path(&mut self, gradient_type: ShapeGradientFillType) {
        let mut attributes = vec![];

        match gradient_type {
            ShapeGradientFillType::Radial => attributes.push(("path", "circle")),
            ShapeGradientFillType::Rectangular => attributes.push(("path", "rect")),
            ShapeGradientFillType::Path => attributes.push(("path", "shape")),
            ShapeGradientFillType::Linear => {}
        }

        self.writer.xml_start_tag("a:path", &attributes);

        // Write the a:fillToRect element.
        self.write_a_fill_to_rect(gradient_type);

        self.writer.xml_end_tag("a:path");

        // Write the a:tileRect element.
        self.write_a_tile_rect(gradient_type);
    }

    // Write the <a:fillToRect> element.
    fn write_a_fill_to_rect(&mut self, gradient_type: ShapeGradientFillType) {
        let mut attributes = vec![];

        match gradient_type {
            ShapeGradientFillType::Path => {
                attributes.push(("l", "50000"));
                attributes.push(("t", "50000"));
                attributes.push(("r", "50000"));
                attributes.push(("b", "50000"));
            }
            _ => {
                attributes.push(("l", "100000"));
                attributes.push(("t", "100000"));
            }
        }

        self.writer.xml_empty_tag("a:fillToRect", &attributes);
    }

    // Write the <a:tileRect> element.
    fn write_a_tile_rect(&mut self, gradient_type: ShapeGradientFillType) {
        let mut attributes = vec![];

        match gradient_type {
            ShapeGradientFillType::Rectangular | ShapeGradientFillType::Radial => {
                attributes.push(("r", "-100000"));
                attributes.push(("b", "-100000"));
            }
            _ => {}
        }

        self.writer.xml_empty_tag("a:tileRect", &attributes);
    }

    // Write the <a:srgbClr> element.
    fn write_color(&mut self, color: Color, transparency: u8) {
        match color {
            Color::Theme(_, _) => {
                let (scheme, lum_mod, lum_off) = color.chart_scheme();
                if !scheme.is_empty() {
                    // Write the a:schemeClr element.
                    self.write_a_scheme_clr(scheme, lum_mod, lum_off, transparency);
                }
            }
            Color::Automatic => {
                let attributes = [("val", "window"), ("lastClr", "FFFFFF")];

                self.writer.xml_empty_tag("a:sysClr", &attributes);
            }
            _ => {
                let attributes = [("val", color.rgb_hex_value())];

                if transparency > 0 {
                    self.writer.xml_start_tag("a:srgbClr", &attributes);

                    // Write the a:alpha element.
                    self.write_a_alpha(transparency);

                    self.writer.xml_end_tag("a:srgbClr");
                } else {
                    self.writer.xml_empty_tag("a:srgbClr", &attributes);
                }
            }
        }
    }

    // Write the <a:schemeClr> element.
    fn write_a_scheme_clr(&mut self, scheme: String, lum_mod: u32, lum_off: u32, transparency: u8) {
        let attributes = [("val", scheme)];

        if lum_mod > 0 || lum_off > 0 || transparency > 0 {
            self.writer.xml_start_tag("a:schemeClr", &attributes);

            if lum_mod > 0 {
                // Write the a:lumMod element.
                self.write_a_lum_mod(lum_mod);
            }

            if lum_off > 0 {
                // Write the a:lumOff element.
                self.write_a_lum_off(lum_off);
            }

            if transparency > 0 {
                // Write the a:alpha element.
                self.write_a_alpha(transparency);
            }

            self.writer.xml_end_tag("a:schemeClr");
        } else {
            self.writer.xml_empty_tag("a:schemeClr", &attributes);
        }
    }

    // Write the <a:lumMod> element.
    fn write_a_lum_mod(&mut self, lum_mod: u32) {
        let attributes = [("val", lum_mod.to_string())];

        self.writer.xml_empty_tag("a:lumMod", &attributes);
    }

    // Write the <a:lumOff> element.
    fn write_a_lum_off(&mut self, lum_off: u32) {
        let attributes = [("val", lum_off.to_string())];

        self.writer.xml_empty_tag("a:lumOff", &attributes);
    }

    // Write the <a:alpha> element.
    fn write_a_alpha(&mut self, transparency: u8) {
        let transparency = u32::from(100 - transparency) * 1000;

        let attributes = [("val", transparency.to_string())];

        self.writer.xml_empty_tag("a:alpha", &attributes);
    }

    // Write the <a:noFill> element.
    fn write_a_no_fill(&mut self) {
        self.writer.xml_empty_tag_only("a:noFill");
    }

    // Write the <a:prstDash> element.
    fn write_a_prst_dash(&mut self, line: &ShapeLine) {
        let attributes = [("val", line.dash_type.to_string())];

        self.writer.xml_empty_tag("a:prstDash", &attributes);
    }

    // Write the default <a:schemeClr> element for textboxes.
    fn write_default_scheme_clr(&mut self, tone: &str, is_line: bool) {
        let mut attributes = vec![];

        attributes.push(("val", tone.to_string()));

        if is_line {
            self.writer.xml_start_tag("a:schemeClr", &attributes);
            self.write_a_shade();
            self.writer.xml_end_tag("a:schemeClr");
        } else {
            self.writer.xml_empty_tag("a:schemeClr", &attributes);
        }
    }

    // Write the <a:shade> element.
    fn write_a_shade(&mut self) {
        let attributes = [("val", "50000")];

        self.writer.xml_empty_tag("a:shade", &attributes);
    }

    // Write the <xdr:style> element.
    fn write_style(&mut self) {
        self.writer.xml_start_tag_only("xdr:style");

        // Write the a:lnRef element.
        self.write_a_ln_ref();

        // Write the a:fillRef element.
        self.write_a_fill_ref();

        // Write the a:effectRef element.
        self.write_a_effect_ref();

        // Write the a:fontRef element.
        self.write_a_font_ref();

        self.writer.xml_end_tag("xdr:style");
    }

    // Write the <a:scrgbClr> element.
    fn write_a_scrgb_clr(&mut self) {
        let attributes = [("r", "0"), ("g", "0"), ("b", "0")];

        self.writer.xml_empty_tag("a:scrgbClr", &attributes);
    }

    // Write the <a:lnRef> element.
    fn write_a_ln_ref(&mut self) {
        let attributes = [("idx", "0")];

        self.writer.xml_start_tag("a:lnRef", &attributes);

        // Write the a:scrgbClr element.
        self.write_a_scrgb_clr();

        self.writer.xml_end_tag("a:lnRef");
    }

    // Write the <a:fillRef> element.
    fn write_a_fill_ref(&mut self) {
        let attributes = [("idx", "0")];

        self.writer.xml_start_tag("a:fillRef", &attributes);

        // Write the a:scrgbClr element.
        self.write_a_scrgb_clr();

        self.writer.xml_end_tag("a:fillRef");
    }

    // Write the <a:effectRef> element.
    fn write_a_effect_ref(&mut self) {
        let attributes = [("idx", "0")];

        self.writer.xml_start_tag("a:effectRef", &attributes);

        // Write the a:scrgbClr element.
        self.write_a_scrgb_clr();

        self.writer.xml_end_tag("a:effectRef");
    }

    // Write the <a:fontRef> element.
    fn write_a_font_ref(&mut self) {
        let attributes = [("idx", "minor")];

        self.writer.xml_start_tag("a:fontRef", &attributes);

        // Write the a:schemeClr element.
        self.write_default_scheme_clr("dk1", false);

        self.writer.xml_end_tag("a:fontRef");
    }

    // Write the <xdr:txBody> element.
    fn write_tx_body(&mut self, drawing_info: &DrawingInfo, shape: &Shape) {
        self.writer.xml_start_tag_only("xdr:txBody");

        // Write the a:bodyPr element.
        self.write_a_body_pr(shape);

        // Write the a:lstStyle element.
        self.write_a_lst_style();

        // Ensure at least one paragraph for empty text.
        let text = if drawing_info.name.is_empty() {
            "\n".to_string()
        } else {
            drawing_info.name.clone()
        };

        for text in text.lines() {
            // Write the a:p element.
            self.write_a_p(text, shape);
        }

        self.writer.xml_end_tag("xdr:txBody");
    }

    // Write the <a:bodyPr> element.
    fn write_a_body_pr(&mut self, shape: &Shape) {
        let mut attributes = vec![];

        match shape.text_options.direction {
            ShapeTextDirection::Horizontal => {}
            ShapeTextDirection::Stacked => attributes.push(("vert", "wordArtVert".to_string())),
            ShapeTextDirection::Rotate90 => attributes.push(("vert", "vert".to_string())),
            ShapeTextDirection::Rotate270 => attributes.push(("vert", "vert270".to_string())),
            ShapeTextDirection::Rotate90EastAsian => {
                attributes.push(("vert", "eaVert".to_string()));
            }
        }

        attributes.push(("wrap", "square".to_string()));
        attributes.push(("rtlCol", "0".to_string()));

        match shape.text_options.vertical_alignment {
            crate::ShapeTextVerticalAlignment::Top => {
                attributes.push(("anchor", "t".to_string()));
            }
            crate::ShapeTextVerticalAlignment::Middle => {
                attributes.push(("anchor", "ctr".to_string()));
                attributes.push(("anchorCtr", "0".to_string()));
            }
            crate::ShapeTextVerticalAlignment::Bottom => {
                attributes.push(("anchor", "b".to_string()));
                attributes.push(("anchorCtr", "0".to_string()));
            }
            crate::ShapeTextVerticalAlignment::TopCentered => {
                attributes.push(("anchor", "t".to_string()));
                attributes.push(("anchorCtr", "1".to_string()));
            }
            crate::ShapeTextVerticalAlignment::MiddleCentered => {
                attributes.push(("anchor", "ctr".to_string()));
                attributes.push(("anchorCtr", "1".to_string()));
            }
            crate::ShapeTextVerticalAlignment::BottomCentered => {
                attributes.push(("anchor", "b".to_string()));
                attributes.push(("anchorCtr", "1".to_string()));
            }
        }

        self.writer.xml_empty_tag("a:bodyPr", &attributes);
    }

    // Write the <a:lstStyle> element.
    fn write_a_lst_style(&mut self) {
        self.writer.xml_empty_tag_only("a:lstStyle");
    }

    // Write the <a:p> element.
    fn write_a_p(&mut self, text: &str, shape: &Shape) {
        let font = &shape.font;
        let has_text_link = shape.text_link.is_some();

        self.writer.xml_start_tag_only("a:p");
        self.write_text_alignment(shape);

        if has_text_link {
            self.write_a_fld();

            if text.is_empty() {
                self.write_font_elements("a:rPr", font);
                self.writer.xml_data_element_only("a:t", " ");
                self.writer.xml_end_tag("a:fld");
                self.write_font_elements("a:endParaRPr", font);
            } else {
                self.write_font_elements("a:rPr", font);
                self.writer.xml_empty_tag_only("a:pPr");
                self.writer.xml_data_element_only("a:t", text);
                self.writer.xml_end_tag("a:fld");
                self.write_font_elements("a:endParaRPr", font);
            }
        } else if text.is_empty() {
            self.write_font_elements("a:endParaRPr", font);
        } else {
            self.writer.xml_start_tag_only("a:r");
            self.write_font_elements("a:rPr", font);
            self.writer.xml_data_element_only("a:t", text);
            self.writer.xml_end_tag("a:r");
        }

        self.writer.xml_end_tag("a:p");
    }

    // Write font sub-elements shared between <a:defRPr> and <a:rPr> elements.
    fn write_font_elements(&mut self, tag: &str, font: &ShapeFont) {
        let mut attributes = vec![("lang", "en-US".to_string())];

        if font.size > 0.0 {
            attributes.push(("sz", font.size.to_string()));
        }

        if font.bold {
            attributes.push(("b", "1".to_string()));
        }

        if font.italic {
            attributes.push(("i", "1".to_string()));
        }
        if font.underline {
            attributes.push(("u", "sng".to_string()));
        }

        if font.has_baseline {
            attributes.push(("baseline", "0".to_string()));
        }

        if font.is_latin() || !font.color.is_auto_or_default() {
            self.writer.xml_start_tag(tag, &attributes);

            if !font.color.is_auto_or_default() {
                self.write_a_solid_fill(font.color, 0);
            }

            if font.is_latin() {
                self.write_a_latin("a:latin", font);
                self.write_a_latin("a:cs", font);
            }

            self.writer.xml_end_tag(tag);
        } else {
            self.writer.xml_empty_tag(tag, &attributes);
        }
    }

    // Write the <a:latin> element.
    fn write_a_latin(&mut self, tag: &str, font: &ShapeFont) {
        let mut attributes = vec![];

        if !font.name.is_empty() {
            attributes.push(("typeface", font.name.to_string()));
        }

        if font.pitch_family > 0 {
            attributes.push(("pitchFamily", font.pitch_family.to_string()));
        }

        if font.character_set > 0 || font.pitch_family > 0 {
            attributes.push(("charset", font.character_set.to_string()));
        }

        self.writer.xml_empty_tag(tag, &attributes);
    }

    // Write the <a:fld> element.
    fn write_a_fld(&mut self) {
        let attributes = [
            ("id", "{B8ADDEFE-BF52-4FD4-8C5D-6B85EF6FF707}"),
            ("type", "TxLink"),
        ];

        self.writer.xml_start_tag("a:fld", &attributes);
    }

    // Write the <a:rPr> element for horizontal text alignment.
    fn write_text_alignment(&mut self, shape: &Shape) {
        match shape.text_options.horizontal_alignment {
            ShapeTextHorizontalAlignment::Default => {}
            ShapeTextHorizontalAlignment::Left => {
                self.writer.xml_empty_tag("a:pPr", &[("algn", "l")]);
            }
            ShapeTextHorizontalAlignment::Center => {
                self.writer.xml_empty_tag("a:pPr", &[("algn", "ctr")]);
            }
            ShapeTextHorizontalAlignment::Right => {
                self.writer.xml_empty_tag("a:pPr", &[("algn", "r")]);
            }
        }
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------
#[derive(Clone)]
pub(crate) struct DrawingCoordinates {
    pub(crate) col: u32,
    pub(crate) row: u32,
    pub(crate) col_offset: f64,
    pub(crate) row_offset: f64,
}

impl Default for DrawingInfo {
    fn default() -> Self {
        let from = DrawingCoordinates {
            col: 0,
            row: 0,
            col_offset: 0.0,
            row_offset: 0.0,
        };

        let to = DrawingCoordinates {
            col: 0,
            row: 0,
            col_offset: 0.0,
            row_offset: 0.0,
        };

        DrawingInfo {
            from,
            to,
            col_absolute: 0,
            row_absolute: 0,
            width: 0.0,
            height: 0.0,
            name: String::new(),
            description: String::new(),
            decorative: false,
            rel_id: 0,
            object_movement: ObjectMovement::MoveButDontSizeWithCells,
            drawing_type: DrawingType::Image,
            url: None,
        }
    }
}

#[derive(Clone)]
pub(crate) struct DrawingInfo {
    pub(crate) from: DrawingCoordinates,
    pub(crate) to: DrawingCoordinates,
    pub(crate) col_absolute: u64,
    pub(crate) row_absolute: u64,
    pub(crate) width: f64,
    pub(crate) height: f64,
    pub(crate) name: String,
    pub(crate) description: String,
    pub(crate) decorative: bool,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) rel_id: u32,
    pub(crate) drawing_type: DrawingType,
    pub(crate) url: Option<Url>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum DrawingType {
    Chart,
    Image,
    Shape,
    Vml,
}

// Trait for object such as Images and Charts that translate to a Drawing object.
pub(crate) trait DrawingObject {
    fn x_offset(&self) -> u32;
    fn y_offset(&self) -> u32;
    fn width_scaled(&self) -> f64;
    fn height_scaled(&self) -> f64;
    fn object_movement(&self) -> ObjectMovement;
    fn name(&self) -> String;
    fn alt_text(&self) -> String;
    fn decorative(&self) -> bool;
    fn drawing_type(&self) -> DrawingType;
}
