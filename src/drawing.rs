// drawing - A module for creating the Excel Drawing.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::{xmlwriter::XMLWriter, ObjectMovement};

pub struct Drawing {
    pub(crate) writer: XMLWriter,
    pub(crate) drawings: Vec<DrawingInfo>,
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
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
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
        let attributes = vec![
            (
                "xmlns:xdr",
                "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing".to_string(),
            ),
            (
                "xmlns:a",
                "http://schemas.openxmlformats.org/drawingml/2006/main".to_string(),
            ),
        ];

        self.writer.xml_start_tag_attr("xdr:wsDr", &attributes);
    }

    // Write the <xdr:twoCellAnchor> element.
    fn write_two_cell_anchor(&mut self, index: u32, drawing_info: &DrawingInfo) {
        let mut attributes = vec![];

        match drawing_info.object_movement {
            ObjectMovement::Default | ObjectMovement::MoveButDontSizeWithCells => {
                attributes.push(("editAs", "oneCell".to_string()))
            }
            ObjectMovement::DontMoveOrSizeWithCells => {
                attributes.push(("editAs", "absolute".to_string()))
            }
            _ => (),
        }

        self.writer
            .xml_start_tag_attr("xdr:twoCellAnchor", &attributes);

        // Write the xdr:from and xdr:to elements
        self.write_from(&drawing_info.from);
        self.write_to(&drawing_info.to);

        // Write the xdr:pic element.
        self.write_pic(index, drawing_info);

        self.writer.xml_empty_tag("xdr:clientData");
        self.writer.xml_end_tag("xdr:twoCellAnchor");
    }

    // Write the <xdr:from> element.
    fn write_from(&mut self, coords: &DrawingCoordinates) {
        self.writer.xml_start_tag("xdr:from");

        self.writer
            .xml_data_element("xdr:col", &coords.col.to_string());
        self.writer
            .xml_data_element("xdr:colOff", &coords.col_offset.to_string());
        self.writer
            .xml_data_element("xdr:row", &coords.row.to_string());
        self.writer
            .xml_data_element("xdr:rowOff", &coords.row_offset.to_string());

        self.writer.xml_end_tag("xdr:from");
    }

    // Write the <xdr:to> element.
    fn write_to(&mut self, coords: &DrawingCoordinates) {
        self.writer.xml_start_tag("xdr:to");

        self.writer
            .xml_data_element("xdr:col", &coords.col.to_string());
        self.writer
            .xml_data_element("xdr:colOff", &coords.col_offset.to_string());
        self.writer
            .xml_data_element("xdr:row", &coords.row.to_string());
        self.writer
            .xml_data_element("xdr:rowOff", &coords.row_offset.to_string());

        self.writer.xml_end_tag("xdr:to");
    }

    // Write the <xdr:pic> element.
    fn write_pic(&mut self, index: u32, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag("xdr:pic");

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
        self.writer.xml_start_tag("xdr:nvPicPr");

        // Write the xdr:cNvPr element.
        self.write_c_nv_pr(index, drawing_info);

        // Write the xdr:cNvPicPr element.
        self.writer.xml_start_tag("xdr:cNvPicPr");
        self.write_a_pic_locks();
        self.writer.xml_end_tag("xdr:cNvPicPr");

        self.writer.xml_end_tag("xdr:nvPicPr");
    }

    // Write the <xdr:cNvPr> element.
    fn write_c_nv_pr(&mut self, index: u32, drawing_info: &DrawingInfo) {
        let id = index + 1;
        let name = format!("Picture {index}");

        let mut attributes = vec![("id", id.to_string()), ("name", name)];

        if !drawing_info.description.is_empty() {
            attributes.push(("descr", drawing_info.description.clone()))
        }

        if drawing_info.decorative {
            self.writer.xml_start_tag_attr("xdr:cNvPr", &attributes);
            self.write_decorative();
            self.writer.xml_end_tag("xdr:cNvPr");
        } else {
            self.writer.xml_empty_tag_attr("xdr:cNvPr", &attributes);
        }
    }

    // Write the decorative sub elements.
    fn write_decorative(&mut self) {
        self.writer.xml_start_tag("a:extLst");

        let attributes = vec![("uri", "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}".to_string())];
        self.writer.xml_start_tag_attr("a:ext", &attributes);

        let attributes = vec![
            (
                "xmlns:a16",
                "http://schemas.microsoft.com/office/drawing/2014/main".to_string(),
            ),
            ("id", "{00000000-0008-0000-0000-000002000000}".to_string()),
        ];
        self.writer
            .xml_empty_tag_attr("a16:creationId", &attributes);

        self.writer.xml_end_tag("a:ext");

        let attributes = vec![("uri", "{C183D7F6-B498-43B3-948B-1728B52AA6E4}".to_string())];
        self.writer.xml_start_tag_attr("a:ext", &attributes);

        let attributes = vec![
            (
                "xmlns:adec",
                "http://schemas.microsoft.com/office/drawing/2017/decorative".to_string(),
            ),
            ("val", "1".to_string()),
        ];
        self.writer
            .xml_empty_tag_attr("adec:decorative", &attributes);

        self.writer.xml_end_tag("a:ext");
        self.writer.xml_end_tag("a:extLst");
    }

    // Write the <a:picLocks> element.
    fn write_a_pic_locks(&mut self) {
        let attributes = vec![("noChangeAspect", "1".to_string())];

        self.writer.xml_empty_tag_attr("a:picLocks", &attributes);
    }

    // Write the <xdr:blipFill> element.
    fn write_blip_fill(&mut self, index: u32) {
        self.writer.xml_start_tag("xdr:blipFill");

        // Write the a:blip element.
        self.write_a_blip(index);

        self.writer.xml_start_tag("a:stretch");
        self.writer.xml_empty_tag("a:fillRect");
        self.writer.xml_end_tag("a:stretch");

        self.writer.xml_end_tag("xdr:blipFill");
    }

    // Write the <a:blip> element.
    fn write_a_blip(&mut self, index: u32) {
        let attributes = vec![
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            ),
            ("r:embed", format!("rId{index}")),
        ];

        self.writer.xml_empty_tag_attr("a:blip", &attributes);
    }

    // Write the <xdr:spPr> element.
    fn write_sp_pr(&mut self, drawing_info: &DrawingInfo) {
        self.writer.xml_start_tag("xdr:spPr");
        self.writer.xml_start_tag("a:xfrm");

        // Write the a:off element.
        self.write_a_off(drawing_info);

        // Write the a:ext element.
        self.write_a_ext(drawing_info);

        self.writer.xml_end_tag("a:xfrm");

        // Write the a:prstGeom element.
        self.write_a_prst_geom();

        self.writer.xml_end_tag("xdr:spPr");
    }

    // Write the <a:off> element.
    fn write_a_off(&mut self, drawing_info: &DrawingInfo) {
        let attributes = vec![
            ("x", drawing_info.col_absolute.to_string()),
            ("y", drawing_info.row_absolute.to_string()),
        ];

        self.writer.xml_empty_tag_attr("a:off", &attributes);
    }

    // Write the <a:ext> element.
    fn write_a_ext(&mut self, drawing_info: &DrawingInfo) {
        let attributes = vec![
            ("cx", drawing_info.width.to_string()),
            ("cy", drawing_info.height.to_string()),
        ];

        self.writer.xml_empty_tag_attr("a:ext", &attributes);
    }

    // Write the <a:prstGeom> element.
    fn write_a_prst_geom(&mut self) {
        let attributes = vec![("prst", "rect".to_string())];

        self.writer.xml_start_tag_attr("a:prstGeom", &attributes);
        self.writer.xml_empty_tag("a:avLst");
        self.writer.xml_end_tag("a:prstGeom");
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

#[derive(Clone)]
pub(crate) struct DrawingInfo {
    pub(crate) from: DrawingCoordinates,
    pub(crate) to: DrawingCoordinates,
    pub(crate) col_absolute: u32,
    pub(crate) row_absolute: u32,
    pub(crate) width: f64,
    pub(crate) height: f64,
    pub(crate) description: String,
    pub(crate) decorative: bool,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) rel_id: u32,
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::drawing::*;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut drawing = Drawing::new();

        let from = DrawingCoordinates {
            col: 2,
            row: 1,
            col_offset: 0.0,
            row_offset: 0.0,
        };

        let to = DrawingCoordinates {
            col: 3,
            row: 6,
            col_offset: 533257.0,
            row_offset: 190357.0,
        };

        let drawing_info = DrawingInfo {
            from,
            to,
            col_absolute: 1219200,
            row_absolute: 190500,
            width: 1142857.0,
            height: 1142857.0,
            description: "rust.png".to_string(),
            decorative: false,
            rel_id: 1,
            object_movement: ObjectMovement::Default,
        };

        drawing.drawings.push(drawing_info);

        drawing.assemble_xml_file();

        let got = drawing.writer.read_to_str();
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <xdr:twoCellAnchor editAs="oneCell">
                    <xdr:from>
                    <xdr:col>2</xdr:col>
                    <xdr:colOff>0</xdr:colOff>
                    <xdr:row>1</xdr:row>
                    <xdr:rowOff>0</xdr:rowOff>
                    </xdr:from>
                    <xdr:to>
                    <xdr:col>3</xdr:col>
                    <xdr:colOff>533257</xdr:colOff>
                    <xdr:row>6</xdr:row>
                    <xdr:rowOff>190357</xdr:rowOff>
                    </xdr:to>
                    <xdr:pic>
                    <xdr:nvPicPr>
                        <xdr:cNvPr id="2" name="Picture 1" descr="rust.png"/>
                        <xdr:cNvPicPr>
                        <a:picLocks noChangeAspect="1"/>
                        </xdr:cNvPicPr>
                    </xdr:nvPicPr>
                    <xdr:blipFill>
                        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
                        <a:stretch>
                        <a:fillRect/>
                        </a:stretch>
                    </xdr:blipFill>
                    <xdr:spPr>
                        <a:xfrm>
                        <a:off x="1219200" y="190500"/>
                        <a:ext cx="1142857" cy="1142857"/>
                        </a:xfrm>
                        <a:prstGeom prst="rect">
                        <a:avLst/>
                        </a:prstGeom>
                    </xdr:spPr>
                    </xdr:pic>
                    <xdr:clientData/>
                </xdr:twoCellAnchor>
                </xdr:wsDr>
                "#,
        );

        assert_eq!(expected, got);
    }
}
