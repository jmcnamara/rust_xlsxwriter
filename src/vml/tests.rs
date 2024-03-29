// Vml unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod theme_tests {

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
        let got = vml_to_vec(got);

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
        let got = vml_to_vec(got);

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
