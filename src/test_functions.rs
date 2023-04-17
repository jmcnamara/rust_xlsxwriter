// Shared test functions for unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use regex::Regex;

// Convert XML string/doc into a vector for comparison testing.
pub(crate) fn xml_to_vec(xml_string: &str) -> Vec<String> {
    lazy_static! {
        static ref ELEMENT_DIVIDES: Regex = Regex::new(r">\s*<").unwrap();
    }

    let mut xml_elements: Vec<String> = Vec::new();
    let tokens: Vec<&str> = ELEMENT_DIVIDES.split(xml_string).collect();

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
    lazy_static! {
        static ref WHITESPACE: Regex = Regex::new(r"\s+").unwrap();
    }

    let mut vml_string = vml_string.replace(['\r', '\n'], "");
    vml_string = WHITESPACE.replace_all(&vml_string, " ").into();

    vml_string = vml_string
        .replace("; ", ";")
        .replace('\'', "\"")
        .replace("<x:Anchor> ", "<x:Anchor>");

    xml_to_vec(&vml_string)
}
