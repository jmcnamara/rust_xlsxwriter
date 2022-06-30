// Shared test functions for unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
// Convert XML string/doc into a vector for comparison testing.
pub(crate) fn xml_to_vec(xml_string: &str) -> Vec<String> {
    let mut xml_elements: Vec<String> = Vec::new();
    let re = regex::Regex::new(r">\s*<").unwrap();
    let tokens: Vec<&str> = re.split(xml_string).collect();

    for token in &tokens {
        let mut element = token.trim().to_string();

        // Add back the removed brackets.
        if !element.starts_with('<') {
            element = format!("<{}", element);
        }
        if !element.ends_with('>') {
            element = format!("{}>", element);
        }

        xml_elements.push(element);
    }
    xml_elements
}
