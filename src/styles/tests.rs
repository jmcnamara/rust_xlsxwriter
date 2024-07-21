// Styles unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod styles_tests {

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
        let dxf_formats = vec![];
        let mut styles = Styles::new(
            &xf_formats,
            &dxf_formats,
            1,
            2,
            1,
            vec![],
            false,
            false,
            false,
        );

        styles.assemble_xml_file();

        let got = styles.writer.read_to_str();
        let got = xml_to_vec(got);

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
