// feature_property_bag - A module for creating the Excel featurePropertyBag.xml
// file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use std::{collections::HashSet, io::Cursor};

use crate::xmlwriter::{
    xml_data_element, xml_declaration, xml_empty_tag, xml_end_tag, xml_start_tag,
};

pub struct FeaturePropertyBag {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) feature_property_bags: HashSet<FeaturePropertyBagTypes>,
}

impl FeaturePropertyBag {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new FeaturePropertyBag struct.
    pub(crate) fn new() -> FeaturePropertyBag {
        let writer = Cursor::new(Vec::with_capacity(2048));

        FeaturePropertyBag {
            writer,
            feature_property_bags: HashSet::new(),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the FeaturePropertyBags element.
        self.write_feature_property_bags();

        // Write the Checkbox bag element.
        self.write_checkbox_bag();

        // Write the XFControls bag element.
        self.write_xf_controls_bag();

        // Write the XFComplement bag element.
        self.write_xf_compliment_bag();

        // Write the XFComplements <bag> element.
        self.write_xf_compliments_bag();

        // Write the DXFComplements <bag> element.
        if self
            .feature_property_bags
            .contains(&FeaturePropertyBagTypes::DXFComplements)
        {
            self.write_dxf_compliments_bag();
        }

        // Close the feature_property_bagProperties tag.
        xml_end_tag(&mut self.writer, "FeaturePropertyBags");
    }

    // Write the <FeaturePropertyBags> element.
    fn write_feature_property_bags(&mut self) {
        let attributes = [(
            "xmlns",
            "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag",
        )];

        xml_start_tag(&mut self.writer, "FeaturePropertyBags", &attributes);
    }

    // Write the Checkbox <bag> element.
    fn write_checkbox_bag(&mut self) {
        let attributes = [("type", "Checkbox")];

        xml_empty_tag(&mut self.writer, "bag", &attributes);
    }

    // Write the XFControls <bag> element.
    fn write_xf_controls_bag(&mut self) {
        let attributes = [("type", "XFControls")];

        xml_start_tag(&mut self.writer, "bag", &attributes);

        // Write the bagId element.
        self.write_bag_id("CellControl", "0");

        xml_end_tag(&mut self.writer, "bag");
    }

    // Write the XFComplement <bag> element.
    fn write_xf_compliment_bag(&mut self) {
        let attributes = [("type", "XFComplement")];

        xml_start_tag(&mut self.writer, "bag", &attributes);

        // Write the bagId element.
        self.write_bag_id("XFControls", "1");

        xml_end_tag(&mut self.writer, "bag");
    }

    // Write the XFComplements <bag> element.
    fn write_xf_compliments_bag(&mut self) {
        let attributes = [
            ("type", "XFComplements"),
            ("extRef", "XFComplementsMapperExtRef"),
        ];

        xml_start_tag(&mut self.writer, "bag", &attributes);
        xml_start_tag(&mut self.writer, "a", &[("k", "MappedFeaturePropertyBags")]);

        self.write_bag_id("", "2");

        xml_end_tag(&mut self.writer, "a");
        xml_end_tag(&mut self.writer, "bag");
    }

    // Write the DXFComplements <bag> element.
    fn write_dxf_compliments_bag(&mut self) {
        let attributes = [
            ("type", "DXFComplements"),
            ("extRef", "DXFComplementsMapperExtRef"),
        ];

        xml_start_tag(&mut self.writer, "bag", &attributes);
        xml_start_tag(&mut self.writer, "a", &[("k", "MappedFeaturePropertyBags")]);

        self.write_bag_id("", "2");

        xml_end_tag(&mut self.writer, "a");
        xml_end_tag(&mut self.writer, "bag");
    }

    // Write the <bagId> element.
    fn write_bag_id(&mut self, key: &str, id: &str) {
        let mut attributes = vec![];

        if !key.is_empty() {
            attributes.push(("k", key.to_string()));
        }

        xml_data_element(&mut self.writer, "bagId", id, &attributes);
    }
}

#[derive(Clone, Copy, Eq, PartialEq, Hash)]
pub(crate) enum FeaturePropertyBagTypes {
    XFComplements,
    DXFComplements,
}
