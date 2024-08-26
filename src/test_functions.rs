// Shared test functions for unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

// Convert XML string/doc into a vector for comparison testing.
pub(crate) fn xml_to_vec(xml_string: &str) -> Vec<String> {
    let mut xml_elements = vec![];
    let lines: Vec<_> = xml_string.trim().lines().map(|line| line.trim()).collect();

    for line in lines {
        for token in line.split("><") {
            let mut element = token.to_string();
            element = element.replace('\r', "");

            // Add back the removed brackets.
            if !element.starts_with('<') {
                element = format!("<{element}");
            }
            if !element.ends_with('>') {
                element = format!("{element}>");
            }

            xml_elements.push(element);
        }
    }

    xml_elements
}

// Convert VML string/doc into a vector for comparison testing. Excel VML tends
// to be less structured than other XML so it needs more massaging.
pub(crate) fn vml_to_vec(vml_string: &str) -> Vec<String> {
    let vml_string = vml_string
        .replace("; ", ";")
        .replace('\'', "\"")
        .replace("<x:Anchor> ", "<x:Anchor>");

    xml_to_vec(&vml_string)
}
