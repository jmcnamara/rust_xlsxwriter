// Test helper functions for integration tests. These functions convert Excel
// xml files into vectors of xml elements to make comparison testing easier.
//
// SPDX-License-Identifier: MIT OR Apache-2.0 Copyright 2022, John McNamara,
// jmcnamara@cpan.org

#[cfg(test)]
use pretty_assertions::assert_eq;
use regex::Regex;
use std::collections::HashMap;
use std::collections::HashSet;
use std::fs;
use std::fs::File;
use std::io::Read;

// Simple test runner struct and methods to create a new xlsx output file and
// compare it with an input xlsx file created by Excel.
#[allow(dead_code)]
pub struct TestRunner<'a> {
    testcase: &'a str,
    unique: &'a str,
    input_filename: String,
    output_filename: String,
    ignore_files: HashSet<&'a str>,
}

impl<'a> TestRunner<'a> {
    #[allow(dead_code)]
    pub fn new(testcase: &'a str) -> TestRunner {
        TestRunner {
            testcase,
            unique: "",
            input_filename: "".to_string(),
            output_filename: "".to_string(),
            ignore_files: HashSet::new(),
        }
    }

    // Set string to add to the default output filename to make it unique so
    // that the multiple tests can be run in parallel.
    #[allow(dead_code)]
    pub fn unique(mut self, unique_string: &'a str) -> TestRunner {
        self.unique = unique_string;
        self
    }

    // Ignore certain xml files within the test xlsx files.
    #[allow(dead_code)]
    pub fn ignore_file(mut self, filename: &'a str) -> TestRunner<'a> {
        self.ignore_files.insert(filename);
        self
    }

    // Ignore the files associated with the formula xl/calcChain.xml.
    #[allow(dead_code)]
    pub fn ignore_calc_chain(mut self) -> TestRunner<'a> {
        self.ignore_files.insert("xl/calcChain.xml");
        self.ignore_files.insert("[Content_Types].xml");
        self.ignore_files.insert("xl/_rels/workbook.xml.rels");
        self
    }

    // Initialize the in/out filenames once other properties have been set.
    #[allow(dead_code)]
    pub fn initialize(mut self) -> TestRunner<'a> {
        self.input_filename = format!("tests/input/{}.xlsx", self.testcase);
        self.output_filename = format!("tests/output/rs_{}_{}.xlsx", self.testcase, self.unique);
        self
    }

    // Getter for the output file name which can be passed to the xlsx file
    // creation function.
    #[allow(dead_code)]
    pub fn output_file(&self) -> &String {
        &self.output_filename
    }

    // Test if the input and output file are equal.
    #[allow(dead_code)]
    pub fn assert_eq(&self) {
        let (exp, got) = compare_xlsx_files(
            &self.input_filename,
            &self.output_filename,
            &self.ignore_files,
        );

        assert_eq!(exp, got);
    }

    // Clean up any the temp output file.
    #[allow(dead_code)]
    pub fn cleanup(&self) {
        fs::remove_file(&self.output_filename).unwrap();
    }
}

// Unzip 2 xlsx files and compare whether they have the same filenames and
// structure. If they are the same then we compare each xml file to ensure that
// files created by rust_xlsxwriter are the same as test files created in Excel.
// Returns two String vectors for comparison testing.
fn compare_xlsx_files(
    exp_file: &str,
    got_file: &str,
    ignore_files: &HashSet<&str>,
) -> (Vec<String>, Vec<String>) {
    // Open the xlsx files.
    let exp_fh = match File::open(exp_file) {
        Ok(fh) => fh,
        Err(err) => {
            return (
                vec![exp_file.to_string(), err.to_string()],
                vec![got_file.to_string()],
            )
        }
    };
    let got_fh = match File::open(got_file) {
        Ok(fh) => fh,
        Err(err) => {
            return (
                vec![exp_file.to_string()],
                vec![got_file.to_string(), err.to_string()],
            )
        }
    };

    // Open the zip structure that comprises an xlsx file.
    let mut exp_zip = match zip::ZipArchive::new(exp_fh) {
        Ok(fh) => fh,
        Err(err) => {
            return (
                vec![exp_file.to_string(), err.to_string()],
                vec![got_file.to_string()],
            )
        }
    };
    let mut got_zip = match zip::ZipArchive::new(got_fh) {
        Ok(fh) => fh,
        Err(err) => {
            return (
                vec![exp_file.to_string()],
                vec![got_file.to_string(), err.to_string()],
            )
        }
    };

    // Iterate through each xml file in the xlsx/zip container and read the
    // xml data as a string.
    let mut exp_filenames = vec![];
    let mut got_filenames = vec![];
    let mut exp_xml: HashMap<String, String> = HashMap::new();
    let mut got_xml: HashMap<String, String> = HashMap::new();

    for i in 0..exp_zip.len() {
        let mut file = match exp_zip.by_index(i) {
            Ok(file) => file,
            Err(err) => {
                return (
                    vec![exp_file.to_string(), err.to_string()],
                    vec![got_file.to_string()],
                )
            }
        };

        // Ignore any test specific files like "xl/calcChain.xml".
        if !ignore_files.contains(file.name()) {
            exp_filenames.push(file.name().to_string());
        }

        let mut xml_data = String::new();
        file.read_to_string(&mut xml_data).unwrap();
        exp_xml.insert(file.name().to_string(), xml_data);
    }

    for i in 0..got_zip.len() {
        let mut file = match got_zip.by_index(i) {
            Ok(file) => file,
            Err(err) => {
                return (
                    vec![exp_file.to_string()],
                    vec![got_file.to_string(), err.to_string()],
                )
            }
        };

        // Ignore any test specific files like "xl/calcChain.xml".
        if !ignore_files.contains(file.name()) {
            got_filenames.push(file.name().to_string());
        }

        let mut xml_data = String::new();
        file.read_to_string(&mut xml_data).unwrap();
        got_xml.insert(file.name().to_string(), xml_data);
    }

    // Sort the xlsx filenames/structure
    exp_filenames.sort();
    got_filenames.sort();

    if exp_filenames != got_filenames {
        return (exp_filenames, got_filenames);
    }

    for filename in exp_filenames {
        let mut exp_xml_string = exp_xml.get(&filename).unwrap().to_string();
        let mut got_xml_string = got_xml.get(&filename).unwrap().to_string();

        // Remove author name and creation date metadata from core.x¦ml file.
        if filename == "docProps/core.xml" {
            // Removed author name from test input files created in Excel.
            exp_xml_string = exp_xml_string.replace("John", "");

            // Remove creation date from core.xml file.
            let re = Regex::new(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z").unwrap();
            exp_xml_string = re.replace_all(&exp_xml_string, "").to_string();
            got_xml_string = re.replace_all(&got_xml_string, "").to_string();
        }

        // Remove workbookView dimensions which are almost always different and
        // calcPr which can have different Excel version ids.
        if filename == "xl/workbook.xml" {
            let re = Regex::new(r"<workbookView[^>]*>").unwrap();
            exp_xml_string = re
                .replace_all(&exp_xml_string, "<workbookView/>")
                .to_string();
            got_xml_string = re
                .replace_all(&got_xml_string, "<workbookView/>")
                .to_string();

            let re = Regex::new(r"<calcPr[^>]*>").unwrap();
            exp_xml_string = re.replace_all(&exp_xml_string, "<calcPr/>").to_string();
            got_xml_string = re.replace_all(&got_xml_string, "<calcPr/>").to_string();
        }

        // Convert the xml strings to vectors for easier comparison.
        let mut exp_xml_vec = xml_to_vec(&exp_xml_string);
        let mut got_xml_vec = xml_to_vec(&got_xml_string);

        // Reorder randomized XML elements in some xlsx xml files to
        // allow comparison testing.
        if filename == "[Content_Types].xml" || filename.ends_with(".rels") {
            exp_xml_vec = sort_xml_file_data(exp_xml_vec);
            got_xml_vec = sort_xml_file_data(got_xml_vec);
        }

        // Add the filename to the xml vector to help identify where
        // differences occurs.
        exp_xml_vec.insert(0, filename.to_string());
        got_xml_vec.insert(0, filename.to_string());

        if exp_xml_vec != got_xml_vec {
            return (exp_xml_vec, got_xml_vec);
        }
    }

    (vec![String::from("Ok")], vec![String::from("Ok")])
}

// Convert XML string/doc into a vector for comparison testing.
fn xml_to_vec(xml_string: &str) -> Vec<String> {
    let mut xml_elements: Vec<String> = Vec::new();
    let re = regex::Regex::new(r">\s*<").unwrap();
    let tokens: Vec<&str> = re.split(xml_string).collect();

    for token in &tokens {
        let mut element = token.trim().to_string();
        element = element.replace("\r", "");

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

// Re-order the elements in an vec of XML elements for comparison purposes. This
// is necessary since Excel can produce the elements of some files, for example
// Content_Types and relationship/.rel files, in a semi-random/hash order.
fn sort_xml_file_data(mut xml_elements: Vec<String>) -> Vec<String> {
    // We don't want to sort the start and end elements.
    let first = xml_elements.remove(0);
    let second = xml_elements.remove(0);
    let last = xml_elements.pop().unwrap();

    // Sort the rest of the elements.
    xml_elements.sort();

    // Add back the start and end elements.
    xml_elements.insert(0, second);
    xml_elements.insert(0, first);
    xml_elements.push(last);

    xml_elements
}
