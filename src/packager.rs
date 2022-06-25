// packager - A library for assembling xml files into an Excel XLSX file.
//
// This module is used in conjunction by rust_xlsxwriter to create an Excel XLSX
// container file.
//
// From Wikipedia: The Open Packaging Conventions (OPC) is a container-file
// technology initially created by Microsoft to store a combination of XML and
// non-XML files that together form a single entity such as an Open XML Paper
// Specification (OpenXPS) document.
// http://en.wikipedia.org/wiki/Open_Packaging_Conventions.
//
// At its simplest an Excel XLSX file contains the following elements::
//
//      ____ [Content_Types].xml
//     |
//     |____ docProps
//     | |____ app.xml
//     | |____ core.xml
//     |
//     |____ xl
//     | |____ workbook.xml
//     | |____ worksheets
//     | | |____ sheet1.xml
//     | |
//     | |____ styles.xml
//     | |
//     | |____ theme
//     | | |____ theme1.xml
//     | |
//     | |_____rels
//     | |____ workbook.xml.rels
//     |
//     |_____rels
//       |____ .rels
//
// The Packager class coordinates the classes that represent the elements of the
// package and writes them into the XLSX file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0 Copyright 2022, John McNamara,
// jmcnamara@cpan.org

use zip::write::FileOptions;
use zip::{DateTime, ZipWriter};

pub struct Packager {
    zip: ZipWriter<std::fs::File>,
    zip_options: FileOptions,
}

impl<'a> Packager {
    // Create a new Packager struct.
    pub fn new(filename: &str) -> Packager {
        let path = std::path::Path::new(filename);
        let file = std::fs::File::create(&path).unwrap();

        let zip = zip::ZipWriter::new(file);

        let zip_options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o600)
            .last_modified_time(DateTime::default())
            .large_file(false);

        Packager { zip, zip_options }
    }

    pub fn create_xlsx(&mut self) {
        self.zip.start_file("app.xml", self.zip_options).unwrap();

        self.zip.finish().unwrap();
    }
}
