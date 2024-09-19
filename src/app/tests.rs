// App unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod app_tests {

    use crate::app::App;
    use crate::test_functions::xml_to_vec;
    use crate::xmlwriter;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble1() {
        let mut app = App::new();

        app.add_heading_pair("Worksheets", 1);
        app.add_part_name("Sheet1");

        app.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&app.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                    <Application>Microsoft Excel</Application>
                    <DocSecurity>0</DocSecurity>
                    <ScaleCrop>false</ScaleCrop>
                    <HeadingPairs>
                        <vt:vector size="2" baseType="variant">
                        <vt:variant>
                            <vt:lpstr>Worksheets</vt:lpstr>
                        </vt:variant>
                        <vt:variant>
                            <vt:i4>1</vt:i4>
                        </vt:variant>
                        </vt:vector>
                    </HeadingPairs>
                    <TitlesOfParts>
                        <vt:vector size="1" baseType="lpstr">
                        <vt:lpstr>Sheet1</vt:lpstr>
                        </vt:vector>
                    </TitlesOfParts>
                    <Company>
                    </Company>
                    <LinksUpToDate>false</LinksUpToDate>
                    <SharedDoc>false</SharedDoc>
                    <HyperlinksChanged>false</HyperlinksChanged>
                    <AppVersion>12.0000</AppVersion>
                </Properties>
                "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        let mut app = App::new();

        app.add_heading_pair("Worksheets", 2);
        app.add_part_name("Sheet1");
        app.add_part_name("Sheet2");

        app.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&app.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                <Application>Microsoft Excel</Application>
                <DocSecurity>0</DocSecurity>
                <ScaleCrop>false</ScaleCrop>
                <HeadingPairs>
                    <vt:vector size="2" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>2</vt:i4>
                    </vt:variant>
                    </vt:vector>
                </HeadingPairs>
                <TitlesOfParts>
                    <vt:vector size="2" baseType="lpstr">
                    <vt:lpstr>Sheet1</vt:lpstr>
                    <vt:lpstr>Sheet2</vt:lpstr>
                    </vt:vector>
                </TitlesOfParts>
                <Company>
                </Company>
                <LinksUpToDate>false</LinksUpToDate>
                <SharedDoc>false</SharedDoc>
                <HyperlinksChanged>false</HyperlinksChanged>
                <AppVersion>12.0000</AppVersion>
                </Properties>
                "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble3() {
        let mut app = App::new();

        app.add_heading_pair("Worksheets", 1);
        app.add_heading_pair("Named Ranges", 1);
        app.add_part_name("Sheet1");
        app.add_part_name("Sheet1!Print_Titles");

        app.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&app.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                <Application>Microsoft Excel</Application>
                <DocSecurity>0</DocSecurity>
                <ScaleCrop>false</ScaleCrop>
                <HeadingPairs>
                    <vt:vector size="4" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                    <vt:variant>
                        <vt:lpstr>Named Ranges</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                    </vt:vector>
                </HeadingPairs>
                <TitlesOfParts>
                    <vt:vector size="2" baseType="lpstr">
                    <vt:lpstr>Sheet1</vt:lpstr>
                    <vt:lpstr>Sheet1!Print_Titles</vt:lpstr>
                    </vt:vector>
                </TitlesOfParts>
                <Company>
                </Company>
                <LinksUpToDate>false</LinksUpToDate>
                <SharedDoc>false</SharedDoc>
                <HyperlinksChanged>false</HyperlinksChanged>
                <AppVersion>12.0000</AppVersion>
                </Properties>
                "#,
        );

        assert_eq!(expected, got);
    }
}
