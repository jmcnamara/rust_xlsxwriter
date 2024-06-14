// Shared test functions for unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::static_regex;

// Convert XML string/doc into a vector for comparison testing.
pub(crate) fn xml_to_vec(xml_string: &str) -> Vec<String> {
    let element_dividers = static_regex!(r">\s*<");

    let mut xml_elements: Vec<String> = Vec::new();
    let tokens: Vec<&str> = element_dividers.split(xml_string).collect();

    for token in &tokens {
        let mut element = token.trim().to_string();
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
    xml_elements
}

// Convert VML string/doc into a vector for comparison testing. Excel VML tends
// to be less structured than other XML so it needs more massaging.
pub(crate) fn vml_to_vec(vml_string: &str) -> Vec<String> {
    let whitespace = static_regex!(r"\s+");

    let mut vml_string = vml_string.replace(['\r', '\n'], "");
    vml_string = whitespace.replace_all(&vml_string, " ").into();

    vml_string = vml_string
        .replace("; ", ";")
        .replace('\'', "\"")
        .replace("<x:Anchor> ", "<x:Anchor>");

    xml_to_vec(&vml_string)
}
