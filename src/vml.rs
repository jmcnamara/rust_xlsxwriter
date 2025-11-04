// vml - A module for creating the Excel Vml.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;

use crate::{drawing::DrawingInfo, ColNum, RowNum};

use crate::xmlwriter::{
    xml_data_element, xml_data_element_only, xml_empty_tag, xml_empty_tag_only, xml_end_tag,
    xml_start_tag, xml_start_tag_only,
};

pub struct Vml {
    pub(crate) comments: Vec<VmlInfo>,
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) buttons: Vec<VmlInfo>,
    pub(crate) header_images: Vec<VmlInfo>,
    pub(crate) data_id: String,
    pub(crate) shape_id: u32,
}

impl Vml {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Vml struct.
    pub fn new() -> Vml {
        let writer = Cursor::new(Vec::with_capacity(2048));

        Vml {
            writer,
            buttons: vec![],
            comments: vec![],
            header_images: vec![],
            data_id: String::new(),
            shape_id: 0,
        }
    }

    // Adjust pixel dimensions from 96 DPI to 72 DPI using the 0.25 rounding
    // method applied by Excel.
    fn vml_dpi_size(dimension: f64) -> f64 {
        (dimension + 0.25).floor() * 72.0 / 96.0
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub fn assemble_xml_file(&mut self) {
        let mut z_index = 0;

        // Write the xml element.
        self.write_xml_namespace();

        // Write the o:shapelayout element.
        self.write_shapelayout();

        if !self.buttons.is_empty() {
            // Write the v:shapetype element.
            self.write_button_shapetype();

            for vml_info in &self.buttons.clone() {
                self.shape_id += 1;
                z_index += 1;

                // Write the v:shape element.
                self.write_button_shape(self.shape_id, z_index, vml_info);
            }
        }

        if !self.comments.is_empty() {
            // Write the v:shapetype element.
            self.write_comment_shapetype();

            for vml_info in &self.comments.clone() {
                self.shape_id += 1;
                z_index += 1;

                // Write the v:shape element.
                self.write_comment_shape(self.shape_id, z_index, vml_info);
            }
        }

        if !self.header_images.is_empty() {
            // Write the v:shapetype element.
            self.write_image_shapetype();

            for (z_index, vml_info) in self.header_images.clone().iter().enumerate() {
                self.shape_id += 1;

                // Write the v:shape element.
                self.write_image_shape(z_index + 1, vml_info);
            }
        }

        // Close the xml tag.
        xml_end_tag(&mut self.writer, "xml");
    }

    // Write the <xml> namespace element.
    fn write_xml_namespace(&mut self) {
        let attributes = [
            ("xmlns:v", "urn:schemas-microsoft-com:vml"),
            ("xmlns:o", "urn:schemas-microsoft-com:office:office"),
            ("xmlns:x", "urn:schemas-microsoft-com:office:excel"),
        ];

        xml_start_tag(&mut self.writer, "xml", &attributes);
    }

    // Write the <o:shapelayout> element for shape layout configuration.
    fn write_shapelayout(&mut self) {
        let attributes = [("v:ext", "edit")];

        xml_start_tag(&mut self.writer, "o:shapelayout", &attributes);

        // Write the o:idmap element.
        self.write_idmap();

        xml_end_tag(&mut self.writer, "o:shapelayout");
    }

    // Write the <o:idmap> element for mapping IDs.
    fn write_idmap(&mut self) {
        let attributes = [
            ("v:ext", "edit".to_string()),
            ("data", self.data_id.clone()),
        ];

        xml_empty_tag(&mut self.writer, "o:idmap", &attributes);
    }

    // Write the <v:shapetype> element for buttons.
    fn write_button_shapetype(&mut self) {
        let attributes = [
            ("id", "_x0000_t201"),
            ("coordsize", "21600,21600"),
            ("o:spt", "201"),
            ("path", "m,l,21600r21600,l21600,xe"),
        ];

        xml_start_tag(&mut self.writer, "v:shapetype", &attributes);

        // Write the v:stroke element.
        self.write_stroke();

        // Write the v:path element.
        self.write_button_path();

        // Write the o:lock element.
        self.write_shapetype_lock();

        xml_end_tag(&mut self.writer, "v:shapetype");
    }

    // Write the <v:shapetype> element for comments.
    fn write_comment_shapetype(&mut self) {
        let attributes = [
            ("id", "_x0000_t202"),
            ("coordsize", "21600,21600"),
            ("o:spt", "202"),
            ("path", "m,l,21600r21600,l21600,xe"),
        ];

        xml_start_tag(&mut self.writer, "v:shapetype", &attributes);

        // Write the v:stroke element.
        self.write_stroke();

        // Write the v:path element.
        self.write_comment_path();

        xml_end_tag(&mut self.writer, "v:shapetype");
    }

    // Write the <v:shapetype> element for header images configuration.
    fn write_image_shapetype(&mut self) {
        let attributes = [
            ("id", "_x0000_t75"),
            ("coordsize", "21600,21600"),
            ("o:spt", "75"),
            ("o:preferrelative", "t"),
            ("path", "m@4@5l@4@11@9@11@9@5xe"),
            ("filled", "f"),
            ("stroked", "f"),
        ];

        xml_start_tag(&mut self.writer, "v:shapetype", &attributes);

        // Write the v:stroke element.
        self.write_stroke();

        // Write the v:formulas element.
        self.write_formulas();

        // Write the v:path element.
        self.write_image_path();

        // Write the o:lock element.
        self.write_aspect_ratio_lock();

        xml_end_tag(&mut self.writer, "v:shapetype");
    }

    // Write the <v:stroke> element for stroke configuration.
    fn write_stroke(&mut self) {
        let attributes = [("joinstyle", "miter")];

        xml_empty_tag(&mut self.writer, "v:stroke", &attributes);
    }

    // Write the <v:formulas> element containing shape formulas.
    fn write_formulas(&mut self) {
        xml_start_tag_only(&mut self.writer, "v:formulas");

        self.write_formula("if lineDrawn pixelLineWidth 0");
        self.write_formula("sum @0 1 0");
        self.write_formula("sum 0 0 @1");
        self.write_formula("prod @2 1 2");
        self.write_formula("prod @3 21600 pixelWidth");
        self.write_formula("prod @3 21600 pixelHeight");
        self.write_formula("sum @0 0 1");
        self.write_formula("prod @6 1 2");
        self.write_formula("prod @7 21600 pixelWidth");
        self.write_formula("sum @8 21600 0");
        self.write_formula("prod @7 21600 pixelHeight");
        self.write_formula("sum @10 21600 0");

        xml_end_tag(&mut self.writer, "v:formulas");
    }
    // Write the <v:f> element for a formula.
    fn write_formula(&mut self, equation: &str) {
        let attributes = [("eqn", equation.to_string())];

        xml_empty_tag(&mut self.writer, "v:f", &attributes);
    }

    // Write the <v:path> element for button shapes.
    fn write_button_path(&mut self) {
        let attributes = [
            ("shadowok", "f"),
            ("o:extrusionok", "f"),
            ("strokeok", "f"),
            ("fillok", "f"),
            ("o:connecttype", "rect"),
        ];

        xml_empty_tag(&mut self.writer, "v:path", &attributes);
    }

    // Write the <v:path> element for header image shapes.
    fn write_image_path(&mut self) {
        let attributes = [
            ("o:extrusionok", "f"),
            ("gradientshapeok", "t"),
            ("o:connecttype", "rect"),
        ];

        xml_empty_tag(&mut self.writer, "v:path", &attributes);
    }

    // Write the <v:path> element for comment shapes.
    fn write_comment_path(&mut self) {
        let attributes = [("gradientshapeok", "t"), ("o:connecttype", "rect")];

        xml_empty_tag(&mut self.writer, "v:path", &attributes);
    }

    // Write the <v:path> element for comments.
    fn write_comment_path2(&mut self) {
        let attributes = [("o:connecttype", "none")];

        xml_empty_tag(&mut self.writer, "v:path", &attributes);
    }

    // Write the <o:lock> element for shape locking.
    fn write_shapetype_lock(&mut self) {
        let attributes = [("v:ext", "edit"), ("shapetype", "t")];

        xml_empty_tag(&mut self.writer, "o:lock", &attributes);
    }

    // Write the <o:lock> element for aspect ratio locking.
    fn write_aspect_ratio_lock(&mut self) {
        let attributes = [("v:ext", "edit"), ("aspectratio", "t")];

        xml_empty_tag(&mut self.writer, "o:lock", &attributes);
    }

    // Write the <v:shape> element for button shapes.
    #[allow(clippy::cast_precision_loss)]
    fn write_button_shape(&mut self, vml_shape_id: u32, z_index: u32, vml_info: &VmlInfo) {
        let top = Self::vml_dpi_size(vml_info.drawing_info.row_absolute as f64);
        let left = Self::vml_dpi_size(vml_info.drawing_info.col_absolute as f64);
        let width = Self::vml_dpi_size(vml_info.drawing_info.width);
        let height = Self::vml_dpi_size(vml_info.drawing_info.height);

        let style = format!(
            "position:absolute;\
             margin-left:{left}pt;\
             margin-top:{top}pt;\
             width:{width}pt;\
             height:{height}pt;\
             z-index:{z_index};\
             mso-wrap-style:tight"
        );

        let shape_id = format!("_x0000_s{vml_shape_id}");

        let mut attributes = vec![("id", shape_id), ("type", "#_x0000_t201".to_string())];

        if !vml_info.alt_text.is_empty() {
            attributes.push(("alt", vml_info.alt_text.clone()));
        }

        attributes.push(("style", style));
        attributes.push(("o:button", "t".to_string()));
        attributes.push(("fillcolor", vml_info.fill_color.clone()));
        attributes.push(("strokecolor", "windowText [64]".to_string()));
        attributes.push(("o:insetmode", "auto".to_string()));

        xml_start_tag(&mut self.writer, "v:shape", &attributes);

        // Write the v:fill element.
        self.write_button_fill();

        // Write the o:lock element.
        self.write_rotation_lock(vml_info);

        // Write the v:textbox element.
        self.write_button_textbox(vml_info);

        // Write the x:ClientData element.
        self.write_button_client_data(vml_info);

        xml_end_tag(&mut self.writer, "v:shape");
    }

    // Write the <v:shape> element for comment shapes.
    #[allow(clippy::cast_precision_loss)]
    fn write_comment_shape(&mut self, vml_shape_id: u32, z_index: u32, vml_info: &VmlInfo) {
        let top = Self::vml_dpi_size(vml_info.drawing_info.row_absolute as f64);
        let left = Self::vml_dpi_size(vml_info.drawing_info.col_absolute as f64);
        let width = Self::vml_dpi_size(vml_info.drawing_info.width);
        let height = Self::vml_dpi_size(vml_info.drawing_info.height);

        let mut style = format!(
            "position:absolute;\
             margin-left:{left}pt;\
             margin-top:{top}pt;\
             width:{width}pt;\
             height:{height}pt;\
             z-index:{z_index};"
        );

        if vml_info.is_visible {
            style += "visibility:visible";
        } else {
            style += "visibility:hidden";
        }

        let shape_id = format!("_x0000_s{vml_shape_id}");

        let mut attributes = vec![("id", shape_id), ("type", "#_x0000_t202".to_string())];

        if !vml_info.alt_text.is_empty() {
            attributes.push(("alt", vml_info.alt_text.clone()));
        }

        attributes.push(("style", style));
        attributes.push(("fillcolor", vml_info.fill_color.clone()));
        attributes.push(("o:insetmode", "auto".to_string()));

        xml_start_tag(&mut self.writer, "v:shape", &attributes);

        // Write the v:fill element.
        self.write_comment_fill();

        // Write the v:shadow element.
        self.write_shadow();

        // Write the v:path element.
        self.write_comment_path2();

        // Write the v:textbox element.
        self.write_comment_textbox();

        // Write the x:ClientData element.
        self.write_comment_client_data(vml_info);

        xml_end_tag(&mut self.writer, "v:shape");
    }

    // Write the <v:shape> element for header image shapes.
    fn write_image_shape(&mut self, z_index: usize, vml_info: &VmlInfo) {
        let width = Self::vml_dpi_size(vml_info.width);
        let height = Self::vml_dpi_size(vml_info.height);

        let style = format!(
            "position:absolute;\
             margin-left:0;\
             margin-top:0;\
             width:{width}pt;\
             height:{height}pt;\
             z-index:{z_index}"
        );

        let shape_id = format!("_x0000_s{}", self.shape_id);

        let attributes = [
            ("id", vml_info.header_position.clone()),
            ("o:spid", shape_id),
            ("type", "#_x0000_t75".to_string()),
            ("style", style),
        ];

        xml_start_tag(&mut self.writer, "v:shape", &attributes);

        // Write the v:imagedata element.
        self.write_imagedata(vml_info);

        // Write the o:lock element.
        self.write_rotation_lock(vml_info);

        xml_end_tag(&mut self.writer, "v:shape");
    }

    // Write the <v:imagedata> element for image data.
    fn write_imagedata(&mut self, vml_info: &VmlInfo) {
        let attributes = [
            ("o:relid", format!("rId{}", vml_info.rel_id)),
            ("o:title", vml_info.text.clone()),
        ];

        xml_empty_tag(&mut self.writer, "v:imagedata", &attributes);
    }

    // Write the <o:lock> element for rotation locking.
    fn write_rotation_lock(&mut self, vml_info: &VmlInfo) {
        let mut attributes = vec![("v:ext", "edit".to_string()), ("rotation", "t".to_string())];

        if vml_info.is_scaled {
            attributes.push(("aspectratio", "f".to_string()));
        }

        xml_empty_tag(&mut self.writer, "o:lock", &attributes);
    }

    // Write the <v:fill> element for button fill properties.
    fn write_button_fill(&mut self) {
        let attributes = [
            ("color2", "buttonFace [67]".to_string()),
            ("o:detectmouseclick", "t".to_string()),
        ];

        xml_empty_tag(&mut self.writer, "v:fill", &attributes);
    }

    // Write the <v:fill> element for comment fill properties.
    fn write_comment_fill(&mut self) {
        let attributes = [("color2", "#ffffe1".to_string())];

        xml_empty_tag(&mut self.writer, "v:fill", &attributes);
    }

    // Write the <v:textbox> element for button text box.
    fn write_button_textbox(&mut self, vml_info: &VmlInfo) {
        let attributes = [("style", "mso-direction-alt:auto"), ("o:singleclick", "f")];

        xml_start_tag(&mut self.writer, "v:textbox", &attributes);

        // Write the div element.
        self.write_button_div(vml_info);

        xml_end_tag(&mut self.writer, "v:textbox");
    }

    // Write the <div> element for button text alignment.
    fn write_button_div(&mut self, vml_info: &VmlInfo) {
        let attributes = [("style", "text-align:center")];

        xml_start_tag(&mut self.writer, "div", &attributes);

        // Write the font element.
        self.write_button_font(vml_info);

        xml_end_tag(&mut self.writer, "div");
    }

    // Write the <font> element for button text font properties.
    fn write_button_font(&mut self, vml_info: &VmlInfo) {
        let attributes = [
            ("face", "Calibri".to_string()),
            ("size", "220".to_string()),
            ("color", "#000000".to_string()),
        ];

        xml_data_element(&mut self.writer, "font", &vml_info.text, &attributes);
    }

    // Write the <x:ClientData> element for button client data.
    fn write_button_client_data(&mut self, vml_info: &VmlInfo) {
        let attributes = [("ObjectType", "Button")];

        xml_start_tag(&mut self.writer, "x:ClientData", &attributes);

        // Write the x:Anchor element.
        self.write_anchor(vml_info);

        // Write the x:PrintObject element.
        self.write_print_object();

        // Write the x:AutoFill element.
        self.write_auto_fill();

        // Write the x:FmlaMacro element.
        self.write_fmla_macro(vml_info);

        // Write the x:TextHAlign element.
        self.write_text_halign();

        // Write the x:TextVAlign element.
        self.write_text_valign();

        xml_end_tag(&mut self.writer, "x:ClientData");
    }

    // Write the <v:textbox> element for comment text box.
    fn write_comment_textbox(&mut self) {
        let attributes = [("style", "mso-direction-alt:auto")];

        xml_start_tag(&mut self.writer, "v:textbox", &attributes);

        // Write the div element.
        self.write_comment_div();

        xml_end_tag(&mut self.writer, "v:textbox");
    }

    // Write the <div> element for comment text alignment.
    fn write_comment_div(&mut self) {
        let attributes = [("style", "text-align:left")];

        xml_start_tag(&mut self.writer, "div", &attributes);

        xml_end_tag(&mut self.writer, "div");
    }

    // Write the <x:ClientData> element for comment client data.
    fn write_comment_client_data(&mut self, vml_info: &VmlInfo) {
        let attributes = [("ObjectType", "Note")];

        xml_start_tag(&mut self.writer, "x:ClientData", &attributes);
        xml_empty_tag_only(&mut self.writer, "x:MoveWithCells");
        xml_empty_tag_only(&mut self.writer, "x:SizeWithCells");

        // Write the x:Anchor element.
        self.write_anchor(vml_info);

        // Write the x:AutoFill element.
        self.write_auto_fill();

        // Write the x:Row element.
        self.write_row(vml_info.row);

        // Write the x:Column element.
        self.write_column(vml_info.col);

        // Write the <x:Visible> element.
        if vml_info.is_visible {
            xml_empty_tag_only(&mut self.writer, "x:Visible");
        }

        xml_end_tag(&mut self.writer, "x:ClientData");
    }

    // Write the <x:Anchor> element for shape anchoring.
    fn write_anchor(&mut self, vml_info: &VmlInfo) {
        let anchor = format!(
            "{}, {}, {}, {}, {}, {}, {}, {}",
            vml_info.drawing_info.from.col,
            vml_info.drawing_info.from.col_offset,
            vml_info.drawing_info.from.row,
            vml_info.drawing_info.from.row_offset,
            vml_info.drawing_info.to.col,
            vml_info.drawing_info.to.col_offset,
            vml_info.drawing_info.to.row,
            vml_info.drawing_info.to.row_offset,
        );

        xml_data_element_only(&mut self.writer, "x:Anchor", &anchor);
    }

    // Write the <x:PrintObject> element for print settings.
    fn write_print_object(&mut self) {
        xml_data_element_only(&mut self.writer, "x:PrintObject", "False");
    }

    // Write the <x:AutoFill> element for autofill settings.
    fn write_auto_fill(&mut self) {
        xml_data_element_only(&mut self.writer, "x:AutoFill", "False");
    }

    // Write the <x:FmlaMacro> element for macro formulas.
    fn write_fmla_macro(&mut self, vml_info: &VmlInfo) {
        xml_data_element_only(&mut self.writer, "x:FmlaMacro", &vml_info.macro_name);
    }

    // Write the <x:TextHAlign> element for horizontal text alignment.
    fn write_text_halign(&mut self) {
        xml_data_element_only(&mut self.writer, "x:TextHAlign", "Center");
    }

    // Write the <x:TextVAlign> element for vertical text alignment.
    fn write_text_valign(&mut self) {
        xml_data_element_only(&mut self.writer, "x:TextVAlign", "Center");
    }

    // Write the <v:shadow> element for shadow properties.
    fn write_shadow(&mut self) {
        let attributes = [("on", "t"), ("color", "black"), ("obscured", "t")];

        xml_empty_tag(&mut self.writer, "v:shadow", &attributes);
    }

    // Write the <x:Row> element for row information.
    fn write_row(&mut self, row: RowNum) {
        xml_data_element_only(&mut self.writer, "x:Row", &row.to_string());
    }

    // Write the <x:Column> element for column information.
    fn write_column(&mut self, col: ColNum) {
        xml_data_element_only(&mut self.writer, "x:Column", &col.to_string());
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------
#[derive(Clone)]
pub(crate) struct VmlInfo {
    pub(crate) row: RowNum,
    pub(crate) col: ColNum,
    pub(crate) width: f64,
    pub(crate) height: f64,
    pub(crate) text: String,
    pub(crate) alt_text: String,
    pub(crate) macro_name: String,
    pub(crate) rel_id: u32,
    pub(crate) header_position: String,
    pub(crate) is_scaled: bool,
    pub(crate) drawing_info: DrawingInfo,
    pub(crate) is_visible: bool,
    pub(crate) fill_color: String,
}

impl Default for VmlInfo {
    fn default() -> Self {
        VmlInfo {
            row: 0,
            col: 0,
            width: 0.0,
            height: 0.0,
            text: String::new(),
            alt_text: String::new(),
            macro_name: String::new(),
            rel_id: 0,
            header_position: String::new(),
            is_scaled: false,
            drawing_info: DrawingInfo::default(),
            is_visible: false,
            fill_color: String::new(),
        }
    }
}
