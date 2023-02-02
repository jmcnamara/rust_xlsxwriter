// vml - A module for creating the Excel Vml.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct Vml {
    pub(crate) writer: XMLWriter,
    pub(crate) header_images: Vec<VmlInfo>,
    pub(crate) data_id: u32,
    pub(crate) shape_id: u32,
}

impl Vml {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Vml struct.
    pub fn new() -> Vml {
        let writer = XMLWriter::new();

        Vml {
            writer,
            header_images: vec![],
            data_id: 0,
            shape_id: 0,
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
        // Write the xml element.
        self.write_xml_namespace();

        // Write the o:shapelayout element.
        self.write_shapelayout();

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
        self.writer.xml_end_tag("xml");
    }

    // Write the <xml> element.
    fn write_xml_namespace(&mut self) {
        let attributes = vec![
            ("xmlns:v", "urn:schemas-microsoft-com:vml".to_string()),
            (
                "xmlns:o",
                "urn:schemas-microsoft-com:office:office".to_string(),
            ),
            (
                "xmlns:x",
                "urn:schemas-microsoft-com:office:excel".to_string(),
            ),
        ];

        self.writer.xml_start_tag_attr("xml", &attributes);
    }

    // Write the <o:shapelayout> element.
    fn write_shapelayout(&mut self) {
        let attributes = vec![("v:ext", "edit".to_string())];

        self.writer.xml_start_tag_attr("o:shapelayout", &attributes);

        // Write the o:idmap element.
        self.write_idmap();

        self.writer.xml_end_tag("o:shapelayout");
    }

    // Write the <o:idmap> element.
    fn write_idmap(&mut self) {
        let attributes = vec![
            ("v:ext", "edit".to_string()),
            ("data", self.data_id.to_string()),
        ];

        self.writer.xml_empty_tag_attr("o:idmap", &attributes);
    }

    // Write the <v:shapetype> element.
    fn write_image_shapetype(&mut self) {
        let attributes = vec![
            ("id", "_x0000_t75".to_string()),
            ("coordsize", "21600,21600".to_string()),
            ("o:spt", "75".to_string()),
            ("o:preferrelative", "t".to_string()),
            ("path", "m@4@5l@4@11@9@11@9@5xe".to_string()),
            ("filled", "f".to_string()),
            ("stroked", "f".to_string()),
        ];

        self.writer.xml_start_tag_attr("v:shapetype", &attributes);

        // Write the v:stroke element.
        self.write_stroke();

        // Write the v:formulas element.
        self.write_formulas();

        // Write the v:path element.
        self.write_path();

        // Write the o:lock element.
        self.write_shapetype_lock();

        self.writer.xml_end_tag("v:shapetype");
    }

    // Write the <v:stroke> element.
    fn write_stroke(&mut self) {
        let attributes = vec![("joinstyle", "miter".to_string())];

        self.writer.xml_empty_tag_attr("v:stroke", &attributes);
    }

    // Write the <v:formulas> element.
    fn write_formulas(&mut self) {
        self.writer.xml_start_tag("v:formulas");

        self.write_formula_with_format("if lineDrawn pixelLineWidth 0");
        self.write_formula_with_format("sum @0 1 0");
        self.write_formula_with_format("sum 0 0 @1");
        self.write_formula_with_format("prod @2 1 2");
        self.write_formula_with_format("prod @3 21600 pixelWidth");
        self.write_formula_with_format("prod @3 21600 pixelHeight");
        self.write_formula_with_format("sum @0 0 1");
        self.write_formula_with_format("prod @6 1 2");
        self.write_formula_with_format("prod @7 21600 pixelWidth");
        self.write_formula_with_format("sum @8 21600 0");
        self.write_formula_with_format("prod @7 21600 pixelHeight");
        self.write_formula_with_format("sum @10 21600 0");

        self.writer.xml_end_tag("v:formulas");
    }
    // Write the <v:f> element.
    fn write_formula_with_format(&mut self, equation: &str) {
        let attributes = vec![("eqn", equation.to_string())];

        self.writer.xml_empty_tag_attr("v:f", &attributes);
    }

    // Write the <v:path> element.
    fn write_path(&mut self) {
        let attributes = vec![
            ("o:extrusionok", "f".to_string()),
            ("gradientshapeok", "t".to_string()),
            ("o:connecttype", "rect".to_string()),
        ];

        self.writer.xml_empty_tag_attr("v:path", &attributes);
    }

    // Write the <o:lock> element.
    fn write_shapetype_lock(&mut self) {
        let attributes = vec![
            ("v:ext", "edit".to_string()),
            ("aspectratio", "t".to_string()),
        ];

        self.writer.xml_empty_tag_attr("o:lock", &attributes);
    }

    // Write the <v:shape> element.
    fn write_image_shape(&mut self, z_index: usize, vml_info: &VmlInfo) {
        let width = vml_info.width;
        let height = vml_info.height;

        let style = format!(
            "position:absolute;\
             margin-left:0;\
             margin-top:0;\
             width:{width}pt;\
             height:{height}pt;\
             z-index:{z_index}"
        );

        let shape_id = format!("_x0000_s{}", self.shape_id);

        let attributes = vec![
            ("id", vml_info.position.to_string()),
            ("o:spid", shape_id),
            ("type", "#_x0000_t75".to_string()),
            ("style", style),
        ];

        self.writer.xml_start_tag_attr("v:shape", &attributes);

        // Write the v:imagedata element.
        self.write_imagedata(vml_info);

        // Write the o:lock element.
        self.write_shape_lock(vml_info);

        self.writer.xml_end_tag("v:shape");
    }

    // Write the <v:imagedata> element.
    fn write_imagedata(&mut self, vml_info: &VmlInfo) {
        let attributes = vec![
            ("o:relid", format!("rId{}", vml_info.rel_id)),
            ("o:title", vml_info.title.to_string()),
        ];

        self.writer.xml_empty_tag_attr("v:imagedata", &attributes);
    }

    // Write the <o:lock> element.
    fn write_shape_lock(&mut self, vml_info: &VmlInfo) {
        let mut attributes = vec![("v:ext", "edit".to_string()), ("rotation", "t".to_string())];

        if vml_info.is_scaled {
            attributes.push(("aspectratio", "f".to_string()));
        }

        self.writer.xml_empty_tag_attr("o:lock", &attributes);
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------
#[derive(Clone)]
pub(crate) struct VmlInfo {
    pub(crate) width: f64,
    pub(crate) height: f64,
    pub(crate) title: String,
    pub(crate) rel_id: u32,
    pub(crate) position: String,
    pub(crate) is_scaled: bool,
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::vml::Vml;
    use crate::{test_functions::vml_to_vec, vml::VmlInfo};

    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble1() {
        let mut vml = Vml::new();

        let vml_info = VmlInfo {
            width: 24.0,
            height: 24.0,
            title: "red".to_string(),
            rel_id: 1,
            position: "LH".to_string(),
            is_scaled: false,
        };

        vml.header_images.push(vml_info);
        vml.data_id = 1;
        vml.shape_id = 1024;

        vml.assemble_xml_file();

        let got = vml.writer.read_to_str();
        let got = vml_to_vec(&got);

        let expected = vml_to_vec(
            r##"
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                <o:shapelayout v:ext="edit">
                  <o:idmap v:ext="edit" data="1"/>
                </o:shapelayout>
                <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                  <v:stroke joinstyle="miter"/>
                  <v:formulas>
                    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                    <v:f eqn="sum @0 1 0"/>
                    <v:f eqn="sum 0 0 @1"/>
                    <v:f eqn="prod @2 1 2"/>
                    <v:f eqn="prod @3 21600 pixelWidth"/>
                    <v:f eqn="prod @3 21600 pixelHeight"/>
                    <v:f eqn="sum @0 0 1"/>
                    <v:f eqn="prod @6 1 2"/>
                    <v:f eqn="prod @7 21600 pixelWidth"/>
                    <v:f eqn="sum @8 21600 0"/>
                    <v:f eqn="prod @7 21600 pixelHeight"/>
                    <v:f eqn="sum @10 21600 0"/>
                  </v:formulas>
                  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                  <o:lock v:ext="edit" aspectratio="t"/>
                </v:shapetype>
                <v:shape id="LH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:24pt;height:24pt;z-index:1">
                  <v:imagedata o:relid="rId1" o:title="red"/>
                  <o:lock v:ext="edit" rotation="t"/>
                </v:shape>
                </xml>
            "##,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        let mut vml = Vml::new();

        let vml_info1 = VmlInfo {
            width: 24.0,
            height: 24.0,
            title: "red".to_string(),
            rel_id: 1,
            position: "LH".to_string(),
            is_scaled: false,
        };

        let vml_info2 = VmlInfo {
            width: 17.25,
            height: 17.25,
            title: "blue".to_string(),
            rel_id: 2,
            position: "CH".to_string(),
            is_scaled: false,
        };

        vml.header_images.push(vml_info1);
        vml.header_images.push(vml_info2);
        vml.data_id = 1;
        vml.shape_id = 1024;

        vml.assemble_xml_file();

        let got = vml.writer.read_to_str();
        let got = vml_to_vec(&got);

        let expected = vml_to_vec(
            r##"
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                <o:shapelayout v:ext="edit">
                    <o:idmap v:ext="edit" data="1"/>
                </o:shapelayout>
                <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                    <v:stroke joinstyle="miter"/>
                    <v:formulas>
                        <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                        <v:f eqn="sum @0 1 0"/>
                        <v:f eqn="sum 0 0 @1"/>
                        <v:f eqn="prod @2 1 2"/>
                        <v:f eqn="prod @3 21600 pixelWidth"/>
                        <v:f eqn="prod @3 21600 pixelHeight"/>
                        <v:f eqn="sum @0 0 1"/>
                        <v:f eqn="prod @6 1 2"/>
                        <v:f eqn="prod @7 21600 pixelWidth"/>
                        <v:f eqn="sum @8 21600 0"/>
                        <v:f eqn="prod @7 21600 pixelHeight"/>
                        <v:f eqn="sum @10 21600 0"/>
                    </v:formulas>
                    <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                    <o:lock v:ext="edit" aspectratio="t"/>
                </v:shapetype>
                <v:shape id="LH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:24pt;height:24pt;z-index:1">
                    <v:imagedata o:relid="rId1" o:title="red"/>
                    <o:lock v:ext="edit" rotation="t"/>
                </v:shape>
                <v:shape id="CH" o:spid="_x0000_s1026" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:17.25pt;height:17.25pt;z-index:2">
                    <v:imagedata o:relid="rId2" o:title="blue"/>
                    <o:lock v:ext="edit" rotation="t"/>
                </v:shape>
                </xml>
            "##,
        );

        assert_eq!(expected, got);
    }
}
