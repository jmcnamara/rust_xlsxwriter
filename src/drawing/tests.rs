// Drawing unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod drawing_tests {

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
            name: "Picture 1".to_string(),
            description: "rust.png".to_string(),
            decorative: false,
            rel_id: 1,
            object_movement: ObjectMovement::MoveButDontSizeWithCells,
            drawing_type: DrawingType::Image,
            url: None,
        };

        drawing.drawings.push(drawing_info);

        drawing.assemble_xml_file();

        let got = drawing.writer.read_to_str();
        let got = xml_to_vec(got);

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
