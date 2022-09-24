// workbook - A module for creating the Excel workbook.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::collections::HashMap;
use std::mem;

use crate::error::XlsxError;
use crate::format::Format;
use crate::packager::Packager;
use crate::packager::PackagerOptions;
use crate::shared_strings_table::SharedStringsTable;
use crate::worksheet::Worksheet;
use crate::xmlwriter::XMLWriter;
use crate::{XlsxColor, XlsxPattern};

/// The workbook struct represents an Excel file in it's entirety. It is the
/// starting point for creating a new Excel xlsx file.
///
/// The Workbook struct represents the entire spreadsheet as you see it in Excel
/// and internally it represents the Excel file as it is written on disk.
///
/// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/demo.png">
///
/// # Examples
///
/// Sample code to generate the Excel file shown above.
///
/// ```rust
/// # // This code is available in examples/app_demo.rs
/// #
/// use chrono::NaiveDate;
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file.
///     let mut workbook = Workbook::new("demo.xlsx");
///
///     // Create some formats to use in the worksheet.
///     let bold_format = Format::new().set_bold();
///     let decimal_format = Format::new().set_num_format("0.000");
///     let date_format = Format::new().set_num_format("yyyy-mm-dd");
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Set the column width for clarity.
///     worksheet.set_column_width(0, 15)?;
///
///     // Write a string without formatting.
///     worksheet.write_string_only(0, 0, "Hello")?;
///
///     // Write a string with the bold format defined above.
///     worksheet.write_string(1, 0, "World", &bold_format)?;
///
///     // Write some numbers.
///     worksheet.write_number_only(2, 0, 1)?;
///     worksheet.write_number_only(3, 0, 2.34)?;
///
///     // Write a number with formatting.
///     worksheet.write_number(4, 0, 3.00, &decimal_format)?;
///
///     // Write a formula.
///     worksheet.write_formula_only(5, 0, "=SIN(PI()/4)")?;
///
///     // Write the date .
///     let date = NaiveDate::from_ymd(2023, 1, 25);
///     worksheet.write_date(6, 0, date, &date_format)?;
///
///     workbook.close()?;
///
///     Ok(())
/// }
/// ```
pub struct Workbook<'a> {
    pub(crate) writer: XMLWriter,
    filename: &'a str,
    worksheets: Vec<Worksheet>,
    xf_formats: Vec<Format>,
    xf_indices: HashMap<String, u32>,
    font_count: u16,
    fill_count: u16,
    border_count: u16,
    num_format_count: u16,
}

impl<'a> Workbook<'a> {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Workbook object to represent an Excel spreadsheet file.
    ///
    /// The `Workbook::new()` constructor is used to create a new Excel workbook
    /// with a given filename.
    ///
    /// # Arguments
    ///
    /// * `filename` - The name of the new Excel file to create. The lifetime of
    ///   the argument lasts for the lifetime of the workbook which is generally
    ///   until the file is written with [`workbook.close()`](Workbook::close).
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a simple workbook, with one
    /// unused worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_new.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new("workbook.xlsx");
    ///
    ///     _ = workbook.add_worksheet();
    ///
    ///     workbook.close()?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/workbook_new.png">
    ///
    pub fn new(filename: &'a str) -> Workbook {
        let writer = XMLWriter::new();
        let default_format = Format::new();
        let xf_indices = HashMap::from([(default_format.format_key(), 0)]);

        Workbook {
            writer,
            filename,
            worksheets: vec![],
            xf_formats: vec![default_format],
            xf_indices,
            font_count: 0,
            fill_count: 0,
            border_count: 0,
            num_format_count: 0,
        }
    }

    /// Add a new worksheet to a workbook.
    ///
    /// The `add_worksheet()` method adds a new [`worksheet`](Worksheet) to a
    /// workbook.
    ///
    /// The worksheets will be given standard Excel name like `Sheet1`,
    /// `Sheet2`, etc. Alternatively, the name can be set using
    /// `worksheet.set_name()`, see the example below and the docs for
    /// [`worksheet.set_name()`](Worksheet::set_name).
    ///
    /// The `add_worksheet()` method returns a mutable borrowed reference to a
    /// Worksheet instance owned by the Workbook so only one worksheet can be in
    /// existence at a time, see the example below. This limitation will be
    /// removed, via other Worksheet creation methods, in an upcoming release.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating adding worksheets to a workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_add_worksheet.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new("workbook.xlsx");
    ///
    ///     let worksheet = workbook.add_worksheet(); // Sheet1
    ///     worksheet.write_string_only(0, 0, "Hello")?;
    ///
    ///     let worksheet = workbook.add_worksheet().set_name("Foglio2")?;
    ///     worksheet.write_string_only(0, 0, "Hello")?;
    ///
    ///     let worksheet = workbook.add_worksheet(); // Sheet3
    ///     worksheet.write_string_only(0, 0, "Hello")?;
    ///
    ///     workbook.close()?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/workbook_add_worksheet.png">
    ///
    pub fn add_worksheet(&mut self) -> &mut Worksheet {
        let sheet_name = format!("Sheet{}", self.worksheets.len() + 1);

        let worksheet = Worksheet::new(sheet_name);
        self.worksheets.push(worksheet);
        let worksheet = self.worksheets.last_mut().unwrap();

        worksheet
    }

    /// Close the Workbook object and write the XLSX file.
    ///
    /// The workbook close() method writes all data to the xlsx file and closes
    /// it.
    ///
    ///
    /// # Errors
    ///
    /// * [`XlsxError::SheetnameReused`] - Worksheet name is already in use in
    ///   the workbook.
    /// * [`XlsxError::IoError`] - A wrapper for various IO errors when creating
    ///   the xlsx file, or its sub-files.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating and closing a simple
    /// workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_new.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new("workbook.xlsx");
    ///
    ///     _ = workbook.add_worksheet();
    ///
    ///     workbook.close()?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn close(&mut self) -> Result<(), XlsxError> {
        // Ensure that there is at least one worksheet in the workbook.
        if self.worksheets.is_empty() {
            self.add_worksheet();
        }
        // Ensure one sheet is selected.
        self.worksheets[0].selected = true;

        // Convert any local formats to workbook/global formats.
        let mut worksheet_formats: Vec<Vec<Format>> = vec![];
        for worksheet in self.worksheets.iter() {
            let formats = worksheet.xf_formats.clone();
            worksheet_formats.push(formats);
        }

        let mut worksheet_indices: Vec<Vec<u32>> = vec![];
        for formats in &mut worksheet_formats {
            let mut indices = vec![];
            for format in formats {
                let index = self.format_index(format);
                indices.push(index);
            }
            worksheet_indices.push(indices);
        }

        for (i, worksheet) in self.worksheets.iter_mut().enumerate() {
            worksheet.set_global_xf_indices(&worksheet_indices[i]);
        }

        // Prepare the formats for writing with styles.rs.
        self.prepare_format_properties();

        // Create the Packager object that will assemble the zip/xlsx file.
        let mut packager = Packager::new(self.filename)?;
        let mut package_options = PackagerOptions::new();

        // Set some packager options, and check for duplicate worksheet names.
        package_options.num_worksheets = self.worksheets.len() as u16;
        for worksheet in self.worksheets.iter() {
            let sheet_name = worksheet.name.clone();

            // Check for duplicate sheet names, which aren't allowed by Excel.
            if package_options.worksheet_names.contains(&sheet_name) {
                return Err(XlsxError::SheetnameReused(sheet_name));
            }

            package_options.worksheet_names.push(sheet_name);

            if worksheet.uses_string_table {
                package_options.has_sst_table = true;
            }
        }

        // Start the zip/xlsx container.
        packager.create_root_files(&package_options);

        // Write the styles.xml file to the zip/xlsx container.

        packager.create_styles_file(
            &self.xf_formats,
            self.font_count,
            self.fill_count,
            self.border_count,
            self.num_format_count,
        );

        // Write the workbook to the zip/xlsx container.
        packager.write_workbook_file(self);

        // Write the worksheets to the zip/xlsx container.
        let mut string_table = SharedStringsTable::new();
        for (index, worksheet) in self.worksheets.iter_mut().enumerate() {
            packager.write_worksheet_file(worksheet, index + 1, &mut string_table);
        }

        // Write the share string table.
        if package_options.has_sst_table {
            packager.write_shared_strings_file(&string_table);
        }

        // Write the docProp files to the zip/xlsx container.
        packager.create_doc_prop_files(&package_options);

        // Close and write the final zip/xlsx container.
        packager.close();

        Ok(())
    }

    // Set the index for the format. This is currently only used in testing but
    // will be used publicly at a later stage.
    #[doc(hidden)]
    pub fn register_format(&mut self, format: &mut Format) {
        let format_key = format.format_key();

        match self.xf_indices.get_mut(&format_key) {
            Some(xf_index) => {
                format.set_xf_index(*xf_index);
            }
            None => {
                let xf_index = self.xf_formats.len() as u32;
                self.xf_formats.push(format.clone());
                format.set_xf_index(xf_index);

                self.xf_indices.insert(format_key, xf_index);
            }
        }
    }

    // -----------------------------------------------------------------------
    // Internal function/methods.
    // -----------------------------------------------------------------------

    // Evaluate and clone formats from worksheets into a workbook level vector
    // of unique format. Also return the index for use in remapping worksheet
    // format indices.
    fn format_index(&mut self, format: &Format) -> u32 {
        let format_key = format.format_key();

        match self.xf_indices.get_mut(&format_key) {
            Some(xf_index) => *xf_index,
            None => {
                let xf_index = self.xf_formats.len() as u32;
                self.xf_formats.push(format.clone());
                self.xf_indices.insert(format_key, xf_index);
                xf_index
            }
        }
    }

    // Prepare all Format properties prior to passing them to styles.rs.
    fn prepare_format_properties(&mut self) {
        // Set the font index for the format objects.
        self.prepare_fonts();

        // Set the fill index for the format objects.
        self.prepare_fills();

        // Set the border index for the format objects.
        self.prepare_borders();

        // Set the number format index for the format objects.
        self.prepare_num_formats();
    }

    // Set the font index for the format objects.
    fn prepare_fonts(&mut self) {
        let mut font_count: u16 = 0;
        let mut font_indices: HashMap<String, u16> = HashMap::new();

        for xf_format in &mut self.xf_formats {
            let font_key = xf_format.font_key();

            match font_indices.get(&font_key) {
                Some(font_index) => {
                    xf_format.set_font_index(*font_index, false);
                }
                None => {
                    font_indices.insert(font_key, font_count);
                    xf_format.set_font_index(font_count, true);
                    font_count += 1;
                }
            }
        }
        self.font_count = font_count;
    }

    // Set the fill index for the format objects.
    fn prepare_fills(&mut self) {
        let mut fill_indices: HashMap<String, u16> = HashMap::new();

        // The user defined fill properties start from 2 since there are 2
        // default fills: patternType="none" and patternType="gray125". The
        // following code adds these 2 default fills.
        let mut fill_count: u16 = 2;

        let temp_format = Format::new();
        let mut fill_key = temp_format.fill_key();
        fill_indices.insert(fill_key, 0);
        fill_key = temp_format
            .set_pattern(crate::XlsxPattern::Gray125)
            .fill_key();
        fill_indices.insert(fill_key, 1);

        for xf_format in &mut self.xf_formats {
            // For a solid fill (pattern == "solid") Excel reverses the role of
            // foreground and background colors, and
            if xf_format.pattern == XlsxPattern::Solid
                && xf_format.background_color != XlsxColor::Automatic
                && xf_format.foreground_color != XlsxColor::Automatic
            {
                mem::swap(
                    &mut xf_format.foreground_color,
                    &mut xf_format.background_color,
                );
            }

            // If the user specifies a foreground or background color without a
            // pattern they probably wanted a solid fill, so we fill in the
            // defaults.
            //
            if (xf_format.pattern == XlsxPattern::None || xf_format.pattern == XlsxPattern::Solid)
                && xf_format.background_color != XlsxColor::Automatic
                && xf_format.foreground_color == XlsxColor::Automatic
            {
                xf_format.foreground_color = xf_format.background_color;
                xf_format.background_color = XlsxColor::Automatic;
                xf_format.pattern = XlsxPattern::Solid;
            }

            if (xf_format.pattern == XlsxPattern::None || xf_format.pattern == XlsxPattern::Solid)
                && xf_format.background_color == XlsxColor::Automatic
                && xf_format.foreground_color != XlsxColor::Automatic
            {
                xf_format.background_color = XlsxColor::Automatic;
                xf_format.pattern = XlsxPattern::Solid;
            }

            // Get a unique fill identifier.
            let fill_key = xf_format.fill_key();

            // Find unique or repeated fill ids.
            match fill_indices.get(&fill_key) {
                Some(fill_index) => {
                    xf_format.set_fill_index(*fill_index, false);
                }
                None => {
                    fill_indices.insert(fill_key, fill_count);
                    xf_format.set_fill_index(fill_count, true);
                    fill_count += 1;
                }
            }
        }
        self.fill_count = fill_count;
    }

    // Set the border index for the format objects.
    fn prepare_borders(&mut self) {
        let mut border_count: u16 = 0;
        let mut border_indices: HashMap<String, u16> = HashMap::new();

        for xf_format in &mut self.xf_formats {
            let border_key = xf_format.border_key();

            match border_indices.get(&border_key) {
                Some(border_index) => {
                    xf_format.set_border_index(*border_index, false);
                }
                None => {
                    border_indices.insert(border_key, border_count);
                    xf_format.set_border_index(border_count, true);
                    border_count += 1;
                }
            }
        }
        self.border_count = border_count;
    }

    // Set the number format index for the format objects.
    fn prepare_num_formats(&mut self) {
        let mut num_formats: HashMap<String, u16> = HashMap::new();
        // User defined number formats in Excel start from index 164.
        let mut index = 164;

        for xf_format in &mut self.xf_formats {
            if xf_format.num_format_index > 0 {
                continue;
            }

            if xf_format.num_format.is_empty() {
                continue;
            }

            let num_format_string = xf_format.num_format.clone();

            match num_formats.get(&num_format_string) {
                Some(index) => {
                    xf_format.set_num_format_index_u16(*index);
                }
                None => {
                    num_formats.insert(num_format_string, index);
                    xf_format.set_num_format_index_u16(index);
                    index += 1;
                    self.num_format_count += 1;
                }
            }
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the workbook element.
        self.write_workbook();

        // Write the fileVersion element.
        self.write_file_version();

        // Write the workbookPr element.
        self.write_workbook_pr();

        // Write the bookViews element.
        self.write_book_views();

        // Write the sheets element.
        self.write_sheets();

        // Write the calcPr element.
        self.write_calc_pr();

        // Close the workbook tag.
        self.writer.xml_end_tag("workbook");
    }

    // Write the <workbook> element.
    fn write_workbook(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let xmlns_r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string();

        let attributes = vec![("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag_attr("workbook", &attributes);
    }

    // Write the <fileVersion> element.
    fn write_file_version(&mut self) {
        let attributes = vec![
            ("appName", "xl".to_string()),
            ("lastEdited", "4".to_string()),
            ("lowestEdited", "4".to_string()),
            ("rupBuild", "4505".to_string()),
        ];

        self.writer.xml_empty_tag_attr("fileVersion", &attributes);
    }

    // Write the <workbookPr> element.
    fn write_workbook_pr(&mut self) {
        let attributes = vec![("defaultThemeVersion", "124226".to_string())];

        self.writer.xml_empty_tag_attr("workbookPr", &attributes);
    }

    // Write the <bookViews> element.
    fn write_book_views(&mut self) {
        self.writer.xml_start_tag("bookViews");

        // Write the workbookView element.
        self.write_workbook_view();

        self.writer.xml_end_tag("bookViews");
    }

    // Write the <workbookView> element.
    fn write_workbook_view(&mut self) {
        let attributes = vec![
            ("xWindow", "240".to_string()),
            ("yWindow", "15".to_string()),
            ("windowWidth", "16095".to_string()),
            ("windowHeight", "9660".to_string()),
        ];

        self.writer.xml_empty_tag_attr("workbookView", &attributes);
    }

    // Write the <sheets> element.
    fn write_sheets(&mut self) {
        self.writer.xml_start_tag("sheets");

        let mut worksheet_names = vec![];
        for worksheet in self.worksheets.iter() {
            worksheet_names.push(worksheet.name.clone());
        }

        for (index, name) in worksheet_names.iter().enumerate() {
            // Write the sheet element.
            self.write_sheet(name, (index + 1) as u16);
        }

        self.writer.xml_end_tag("sheets");
    }

    // Write the <sheet> element.
    fn write_sheet(&mut self, name: &str, index: u16) {
        //let name = name;
        let sheet_id = format!("{}", index);
        let ref_id = format!("rId{}", index);

        let attributes = vec![
            ("name", name.to_string()),
            ("sheetId", sheet_id),
            ("r:id", ref_id),
        ];

        self.writer.xml_empty_tag_attr("sheet", &attributes);
    }

    // Write the <calcPr> element.
    fn write_calc_pr(&mut self) {
        let attributes = vec![
            ("calcId", "124519".to_string()),
            ("fullCalcOnLoad", "1".to_string()),
        ];

        self.writer.xml_empty_tag_attr("calcPr", &attributes);
    }
}

#[cfg(test)]
mod tests {

    use super::Workbook;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut workbook = Workbook::new("test.xlsx");
        workbook.add_worksheet();

        workbook.assemble_xml_file();

        let got = workbook.writer.read_to_string();
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
              <workbookPr defaultThemeVersion="124226"/>
              <bookViews>
                <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
              </bookViews>
              <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
              </sheets>
              <calcPr calcId="124519" fullCalcOnLoad="1"/>
            </workbook>
            "#,
        );

        assert_eq!(got, expected);
    }
}
