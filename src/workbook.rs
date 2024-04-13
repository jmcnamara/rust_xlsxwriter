// workbook - A module for creating the Excel workbook.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! # Working with Workbooks
//!
//! The [`Workbook`] struct represents an Excel file in it's entirety. It is the
//! starting point for creating a new Excel xlsx file.
//!
//!
//! ```
//! # // This code is available in examples/doc_workbook_new.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     let _worksheet = workbook.add_worksheet();
//!
//!     workbook.save("workbook.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/workbook_new.png">
//!
//!
//! For more details on the Worksheet APIs for see the [`Worksheet`]
//! documentation and the sections below.
//!
//! # Contents
//!
//! - [Creating and saving an xlsx file](#creating-and-saving-an-xlsx-file)
//! - [Checksum of a saved file](#checksum-of-a-saved-file)
//!
//!
//! # Creating and saving an xlsx file
//!
//! Creating a  [`Workbook`] struct instance to represent an Excel xlsx file is
//! done via the [`Workbook::new()`] method:
//!
//!
//! ```
//! # // This code is available in examples/doc_workbook_new.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//! #     let _worksheet = workbook.add_worksheet();
//! #
//! #     workbook.save("workbook.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Once you are finished writing data via a worksheet you can save it with the
//! [`Workbook::save()`] method:
//!
//! ```
//! # // This code is available in examples/doc_workbook_new.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     let _worksheet = workbook.add_worksheet();
//!
//!     workbook.save("workbook.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! This will you a simple output file like the following.
//!
//! <img src="https://rustxlsxwriter.github.io/images/workbook_new.png">
//!
//! The  `save()` method takes a [`std::path`] or path/filename string. You can
//! also save the xlsx file data to a `Vec<u8>` buffer via the
//! [`Workbook::save_to_buffer()`] method:
//!
//! ```
//! # // This code is available in examples/doc_workbook_save_to_buffer.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     let worksheet = workbook.add_worksheet();
//!     worksheet.write_string(0, 0, "Hello")?;
//!
//!     let buf = workbook.save_to_buffer()?;
//!
//!     println!("File size: {}", buf.len());
//!
//!     Ok(())
//! }
//! ```
//!
//! This can be useful if you intend to stream the data.
//!
//!
//! # Checksum of a saved file
//!
//!
//! A common issue that occurs with `rust_xlsxwriter`, but also with Excel, is
//! that running the same program twice doesn't generate the same file, byte for
//! byte. This can cause issues with applications that do checksumming for
//! testing purposes.
//!
//! For example consider the following simple `rust_xlsxwriter` program:
//!
//! ```
//! # // This code is available in examples/doc_properties_checksum1.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!     let worksheet = workbook.add_worksheet();
//!
//!     worksheet.write_string(0, 0, "Hello")?;
//!
//!     workbook.save("properties.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! If we run this several times, with a small delay, we will get different
//! checksums as shown below:
//!
//! ```bash
//! $ cargo run --example doc_properties_checksum1
//!
//! $ sum properties.xlsx
//! 62457 6 properties.xlsx
//!
//! $ sleep 2
//!
//! $ cargo run --example doc_properties_checksum1
//!
//! $ sum properties.xlsx
//! 56692 6 properties.xlsx # Different to previous.
//! ```
//!
//! This is due to a file creation datetime that is included in the file and
//! which changes each time a new file is created.
//!
//! The relevant section of the `docProps/core.xml` sub-file in the xlsx format
//! looks like this:
//!
//! ```xml
//! <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//! <cp:coreProperties>
//!   <dc:creator/>
//!   <cp:lastModifiedBy/>
//!   <dcterms:created xsi:type="dcterms:W3CDTF">2023-01-08T00:23:58Z</dcterms:created>
//!   <dcterms:modified xsi:type="dcterms:W3CDTF">2023-01-08T00:23:58Z</dcterms:modified>
//! </cp:coreProperties>
//! ```
//!
//! If required this can be avoided by setting a constant creation date in the
//! document properties metadata:
//!
//!
//! ```
//! # // This code is available in examples/doc_properties_checksum2.rs
//! #
//! use rust_xlsxwriter::{DocProperties, ExcelDateTime, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Create a file creation date for the file.
//!     let date = ExcelDateTime::from_ymd(2023, 1, 1)?;
//!
//!     // Add it to the document metadata.
//!     let properties = DocProperties::new().set_creation_datetime(&date);
//!     workbook.set_properties(&properties);
//!
//!     let worksheet = workbook.add_worksheet();
//!     worksheet.write_string(0, 0, "Hello")?;
//!
//!     workbook.save("properties.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Then we will get the same checksum for the same output every time:
//!
//! ```bash
//! $ cargo run --example doc_properties_checksum2
//!
//! $ sum properties.xlsx 8914 6 properties.xlsx
//!
//! $ sleep 2
//!
//! $ cargo run --example doc_properties_checksum2
//!
//! $ sum properties.xlsx 8914 6 properties.xlsx # Same as previous
//! ```
//!
//! For more details see [`DocProperties`] and [`Workbook::set_properties()`].
//!
#![warn(missing_docs)]

mod tests;

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Seek, Write};
use std::mem;
use std::path::Path;

use crate::error::XlsxError;
use crate::format::Format;
use crate::packager::Packager;
use crate::packager::PackagerOptions;
use crate::worksheet::Worksheet;
use crate::xmlwriter::XMLWriter;
use crate::{
    utility, Border, Chart, ChartRange, ChartRangeCacheData, ColNum, DefinedName, DefinedNameType,
    DocProperties, Fill, Font, Image, RowNum, Visible, NUM_IMAGE_FORMATS,
};
use crate::{Color, FormatPattern};

/// The `Workbook` struct represents an Excel file in its entirety. It is the
/// starting point for creating a new Excel xlsx file.
///
/// <img src="https://rustxlsxwriter.github.io/images/demo.png">
///
/// # Examples
///
/// Sample code to generate the Excel file shown above.
///
/// ```rust
/// # // This code is available in examples/app_demo.rs
/// #
/// use rust_xlsxwriter::*;
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Create some formats to use in the worksheet.
///     let bold_format = Format::new().set_bold();
///     let decimal_format = Format::new().set_num_format("0.000");
///     let date_format = Format::new().set_num_format("yyyy-mm-dd");
///     let merge_format = Format::new()
///         .set_border(FormatBorder::Thin)
///         .set_align(FormatAlign::Center);
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Set the column width for clarity.
///     worksheet.set_column_width(0, 22)?;
///
///     // Write a string without formatting.
///     worksheet.write(0, 0, "Hello")?;
///
///     // Write a string with the bold format defined above.
///     worksheet.write_with_format(1, 0, "World", &bold_format)?;
///
///     // Write some numbers.
///     worksheet.write(2, 0, 1)?;
///     worksheet.write(3, 0, 2.34)?;
///
///     // Write a number with formatting.
///     worksheet.write_with_format(4, 0, 3.00, &decimal_format)?;
///
///     // Write a formula.
///     worksheet.write(5, 0, Formula::new("=SIN(PI()/4)"))?;
///
///     // Write a date.
///     let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
///     worksheet.write_with_format(6, 0, &date, &date_format)?;
///
///     // Write some links.
///     worksheet.write(7, 0, Url::new("https://www.rust-lang.org"))?;
///     worksheet.write(8, 0, Url::new("https://www.rust-lang.org").set_text("Rust"))?;
///
///     // Write some merged cells.
///     worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;
///
///     // Insert an image.
///     let image = Image::new("examples/rust_logo.png")?;
///     worksheet.insert_image(1, 2, &image)?;
///
///     // Save the file to disk.
///     workbook.save("demo.xlsx")?;
///
///     Ok(())
/// }
/// ```
pub struct Workbook {
    pub(crate) writer: XMLWriter,
    pub(crate) properties: DocProperties,
    pub(crate) worksheets: Vec<Worksheet>,
    pub(crate) xf_formats: Vec<Format>,
    pub(crate) dxf_formats: Vec<Format>,
    pub(crate) font_count: u16,
    pub(crate) fill_count: u16,
    pub(crate) border_count: u16,
    pub(crate) num_formats: Vec<String>,
    pub(crate) has_hyperlink_style: bool,
    pub(crate) embedded_images: Vec<Image>,
    xf_indices: HashMap<Format, u32>,
    dxf_indices: HashMap<Format, u32>,
    active_tab: u16,
    first_sheet: u16,
    defined_names: Vec<DefinedName>,
    user_defined_names: Vec<DefinedName>,
    read_only_mode: u8,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Workbook object to represent an Excel spreadsheet file.
    ///
    /// The `Workbook::new()` constructor is used to create a new Excel workbook
    /// object. This is used to create worksheets and add data prior to saving
    /// everything to an xlsx file with [`save()`](Workbook::save), or
    /// [`save_to_buffer()`](Workbook::save_to_buffer).
    ///
    /// **Note**: `rust_xlsxwriter` can only create new files. It cannot read or
    /// modify existing files.
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
    /// #
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let _worksheet = workbook.add_worksheet();
    ///
    ///     workbook.save("workbook.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/workbook_new.png">
    ///
    pub fn new() -> Workbook {
        let writer = XMLWriter::new();

        let mut workbook = Workbook {
            writer,
            properties: DocProperties::new(),
            font_count: 0,
            active_tab: 0,
            fill_count: 0,
            first_sheet: 0,
            border_count: 0,
            num_formats: vec![],
            read_only_mode: 0,
            has_hyperlink_style: false,
            worksheets: vec![],
            xf_formats: vec![],
            dxf_formats: vec![],
            defined_names: vec![],
            user_defined_names: vec![],
            xf_indices: HashMap::new(),
            dxf_indices: HashMap::new(),
            embedded_images: vec![],
        };

        // Initialize the workbook with the same function used to reset it.
        Self::reset(&mut workbook);

        workbook
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
    /// The `add_worksheet()` method returns a borrowed mutable reference to a
    /// Worksheet instance owned by the Workbook so only one worksheet can be in
    /// existence at a time, see the example below. This limitation can be
    /// avoided, if necessary, by creating standalone Worksheet objects via
    /// [`Worksheet::new()`] and then later adding them to the workbook with
    /// [`workbook.push_worksheet`](Workbook::push_worksheet).
    ///
    /// See also the documentation on [Creating worksheets] and working with the
    /// borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating adding worksheets to a
    /// workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_add_worksheet.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let worksheet = workbook.add_worksheet(); // Sheet1
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     let worksheet = workbook.add_worksheet().set_name("Foglio2")?;
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     let worksheet = workbook.add_worksheet(); // Sheet3
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     workbook.save("workbook.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/workbook_add_worksheet.png">
    ///
    pub fn add_worksheet(&mut self) -> &mut Worksheet {
        let name = format!("Sheet{}", self.worksheets.len() + 1);

        let mut worksheet = Worksheet::new();
        worksheet.set_name(&name).unwrap();

        self.worksheets.push(worksheet);
        let worksheet = self.worksheets.last_mut().unwrap();

        worksheet
    }

    /// Get a worksheet reference by index.
    ///
    /// Get a reference to a worksheet created via
    /// [`workbook.add_worksheet()`](Workbook::add_worksheet) using an index
    /// based on the creation order.
    ///
    /// Due to borrow checking rules you can only have one active reference to a
    /// worksheet object created by `add_worksheet()` since that method always
    /// returns a mutable reference. For a workbook with multiple worksheets
    /// this restriction is generally workable if you can create and use the
    /// worksheets sequentially since you will only need to have one reference
    /// at any one time. However, if you can't structure your code to work
    /// sequentially then you get a reference to a previously created worksheet
    /// using `worksheet_from_index()`. The standard borrow checking rules still
    /// apply so you will have to give up ownership of any other worksheet
    /// reference prior to calling this method. See the example below.
    ///
    /// See also [`worksheet_from_name()`](Workbook::worksheet_from_name) and
    /// the documentation on [Creating worksheets] and working with the borrow
    /// checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Parameters
    ///
    /// * `index` - The index of the worksheet to get a reference to.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::UnknownWorksheetNameOrIndex`] - Error when trying to
    ///   retrieve a worksheet reference by index. This is usually an index out
    ///   of bounds error.
    ///
    /// # Examples
    ///
    /// The following example demonstrates getting worksheet reference by index.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_worksheet_from_index.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Start with a reference to worksheet1.
    ///     let mut worksheet1 = workbook.add_worksheet();
    ///     worksheet1.write_string(0, 0, "Hello")?;
    ///
    ///     // If we don't try to use the workbook1 reference again we can switch to
    ///     // using a reference to worksheet2.
    ///     let mut worksheet2 = workbook.add_worksheet();
    ///     worksheet2.write_string(0, 0, "Hello")?;
    ///
    ///     // Stop using worksheet2 and move back to worksheet1.
    ///     worksheet1 = workbook.worksheet_from_index(0)?;
    ///     worksheet1.write_string(1, 0, "Sheet1")?;
    ///
    ///     // Stop using worksheet1 and move back to worksheet2.
    ///     worksheet2 = workbook.worksheet_from_index(1)?;
    ///     worksheet2.write_string(1, 0, "Sheet2")?;
    ///
    /// #     workbook.save("workbook.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/workbook_worksheet_from_index.png">
    ///
    pub fn worksheet_from_index(&mut self, index: usize) -> Result<&mut Worksheet, XlsxError> {
        if let Some(worksheet) = self.worksheets.get_mut(index) {
            Ok(worksheet)
        } else {
            Err(XlsxError::UnknownWorksheetNameOrIndex(index.to_string()))
        }
    }

    /// Get a worksheet reference by name.
    ///
    /// Get a reference to a worksheet created via
    /// [`workbook.add_worksheet()`](Workbook::add_worksheet) using the sheet
    /// name.
    ///
    /// Due to borrow checking rules you can only have one active reference to a
    /// worksheet object created by `add_worksheet()` since that method always
    /// returns a mutable reference. For a workbook with multiple worksheets
    /// this restriction is generally workable if you can create and use the
    /// worksheets sequentially since you will only need to have one reference
    /// at any one time. However, if you can't structure your code to work
    /// sequentially then you get a reference to a previously created worksheet
    /// using `worksheet_from_name()`. The standard borrow checking rules still
    /// apply so you will have to give up ownership of any other worksheet
    /// reference prior to calling this method. See the example below.
    ///
    /// Worksheet names are usually "Sheet1", "Sheet2", etc., or else a user
    /// define name that was set using
    /// [`worksheet.set_name()`](Worksheet::set_name). You can also use the
    /// [`worksheet.name()`](Worksheet::name) method to get the name.
    ///
    /// See also [`worksheet_from_index()`](Workbook::worksheet_from_index) and
    /// the documentation on [Creating worksheets] and working with the borrow
    /// checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Parameters
    ///
    /// * `name` - The name of the worksheet to get a reference to.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::UnknownWorksheetNameOrIndex`] - Error when trying to
    ///   retrieve a worksheet reference by index. This is usually an index out
    ///   of bounds error.
    ///
    /// # Examples
    ///
    /// The following example demonstrates getting worksheet reference by index.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_worksheet_from_index.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Start with a reference to worksheet1.
    ///     let mut worksheet1 = workbook.add_worksheet();
    ///     worksheet1.write_string(0, 0, "Hello")?;
    ///
    ///     // If we don't try to use the workbook1 reference again we can switch to
    ///     // using a reference to worksheet2.
    ///     let mut worksheet2 = workbook.add_worksheet();
    ///     worksheet2.write_string(0, 0, "Hello")?;
    ///
    ///     // Stop using worksheet2 and move back to worksheet1.
    ///     worksheet1 = workbook.worksheet_from_index(0)?;
    ///     worksheet1.write_string(1, 0, "Sheet1")?;
    ///
    ///     // Stop using worksheet1 and move back to worksheet2.
    ///     worksheet2 = workbook.worksheet_from_index(1)?;
    ///     worksheet2.write_string(1, 0, "Sheet2")?;
    ///
    /// #     workbook.save("workbook.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/workbook_worksheet_from_name.png">
    ///
    pub fn worksheet_from_name(&mut self, sheetname: &str) -> Result<&mut Worksheet, XlsxError> {
        for (index, worksheet) in self.worksheets.iter_mut().enumerate() {
            if sheetname == worksheet.name {
                return self.worksheet_from_index(index);
            }
        }

        // If we didn't find a matching sheet name then raise
        Err(XlsxError::UnknownWorksheetNameOrIndex(
            sheetname.to_string(),
        ))
    }

    /// Get a mutable reference to the vector of worksheets.
    ///
    /// Get a mutable reference to the vector of Worksheets used by the Workbook
    /// instance. This can be useful for iterating over, and performing the same
    /// operation, on all the worksheets in the workbook. See the example below.
    ///
    /// If you are careful you can also use some of the standard [slice]
    /// operations on the vector reference, see below.
    ///
    /// See also the documentation on [Creating worksheets] and working with the
    /// borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Examples
    ///
    /// The following example demonstrates operating on the vector of all the
    /// worksheets in a workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_worksheets_mut.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Add three worksheets to the workbook.
    ///     let _ = workbook.add_worksheet();
    ///     let _ = workbook.add_worksheet();
    ///     let _ = workbook.add_worksheet();
    ///
    ///     // Write the same data to all three worksheets.
    ///     for worksheet in workbook.worksheets_mut() {
    ///         worksheet.write_string(0, 0, "Hello")?;
    ///         worksheet.write_number(1, 0, 12345)?;
    ///     }
    ///
    ///     // If you are careful you can use standard slice operations.
    ///     workbook.worksheets_mut().swap(0, 1);
    /// #
    /// #     workbook.save("workbook.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file, note the same data is in all three worksheets and Sheet2
    /// and Sheet1 have swapped position:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/workbook_worksheets_mut.png">
    ///
    pub fn worksheets_mut(&mut self) -> &mut Vec<Worksheet> {
        &mut self.worksheets
    }

    /// Get a reference to the vector of worksheets.
    ///
    /// Get a reference to the vector of Worksheets used by the Workbook
    /// instance. This is less useful than
    /// [`worksheets_mut`](Workbook::worksheets_mut) version since a mutable
    /// reference is required for most worksheet operations.
    ///
    /// # Examples
    ///
    /// The following example demonstrates operating on the vector of all the
    /// worksheets in a workbook. The non mutable version of this method is less
    /// useful than `workbook.worksheets_mut()`.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_worksheets.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Add three worksheets to the workbook.
    ///     let _worksheet1 = workbook.add_worksheet();
    ///     let _worksheet2 = workbook.add_worksheet();
    ///     let _worksheet3 = workbook.add_worksheet();
    ///
    ///     // Get some information from all three worksheets.
    ///     for worksheet in workbook.worksheets() {
    ///         println!("{}", worksheet.name());
    ///     }
    ///
    /// #     workbook.save("workbook.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn worksheets(&mut self) -> &Vec<Worksheet> {
        &self.worksheets
    }

    /// Add a worksheet object to a workbook.
    ///
    /// Add a worksheet created directly using `Workbook::new()` to a workbook.
    ///
    /// There are two way of creating a worksheet object with `rust_xlsxwriter`:
    /// via the [`workbook.add_worksheet()`](Workbook::add_worksheet) method and
    /// via the [`Worksheet::new()`] constructor. The first method ties the
    /// worksheet to the workbook object that will write it automatically when
    /// the file is saved, whereas the second method creates a worksheet that is
    /// independent of a workbook. This has certain advantages in keeping the
    /// worksheet free of the workbook borrow checking until you wish to add it.
    ///
    /// When working with the independent worksheet object you can add it to a
    /// workbook using `push_worksheet()`, see the example below.
    ///
    /// See also the documentation on [Creating worksheets] and working with the
    /// borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Parameters
    ///
    /// * `worksheet` - The worksheet to add to the workbook.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a standalone worksheet
    /// object and then adding to a workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_push_worksheet.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    ///     let mut worksheet = Worksheet::new();
    ///
    ///     // Use the worksheet object.
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     // Add it to the workbook.
    ///     workbook.push_worksheet(worksheet);
    ///
    ///     // Save the workbook.
    /// #     workbook.save("workbook.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/workbook_push_worksheet.png">
    ///
    pub fn push_worksheet(&mut self, mut worksheet: Worksheet) {
        if worksheet.name().is_empty() {
            let name = format!("Sheet{}", self.worksheets.len() + 1);
            worksheet.set_name(&name).unwrap();
        }

        self.worksheets.push(worksheet);
    }

    /// Save the Workbook as an xlsx file.
    ///
    /// The workbook `save()` method writes all the Workbook data to a new xlsx
    /// file. It will overwrite any existing file.
    ///
    /// The `save()` method can be called multiple times so it is possible to
    /// get incremental files at different stages of a process, or to save the
    /// same Workbook object to different paths. However, `save()` is an
    /// expensive operation which assembles multiple files into an xlsx/zip
    /// container so for performance reasons you shouldn't call it
    /// unnecessarily.
    ///
    /// # Parameters
    ///
    /// * `path` - The path of the new Excel file to create as a `&str` or as a
    ///   [`std::path`] `Path` or `PathBuf` instance.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::SheetnameReused`] - Worksheet name is already in use in
    ///   the workbook.
    /// * [`XlsxError::TableNameReused`] - Worksheet Table name is already in
    ///   use in the workbook.
    /// * [`XlsxError::IoError`] - A wrapper for various IO errors when creating
    ///   the xlsx file, or its sub-files.
    /// * [`XlsxError::ZipError`] - A wrapper for various zip errors when
    ///   creating the xlsx file, or its sub-files.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a simple workbook using a
    /// string path.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let _worksheet = workbook.add_worksheet();
    ///
    ///     workbook.save("workbook.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// The following example demonstrates creating a simple workbook using a
    /// Rust [`std::path`] Path.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save_to_path.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let _worksheet = workbook.add_worksheet();
    ///
    ///     let path = std::path::Path::new("workbook.xlsx");
    ///     workbook.save(path)?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<(), XlsxError> {
        #[cfg(feature = "test-resave")]
        {
            // Some test code to test double/multiple saves.
            let file = std::fs::File::create(<&std::path::Path>::clone(&path.as_ref()))?;
            self.save_internal(file)?;
        }

        let file = std::fs::File::create(path)?;
        self.save_internal(file)?;
        Ok(())
    }

    /// Save the Workbook as an xlsx file and return it as a byte vector.
    ///
    /// The workbook `save_to_buffer()` method is similar to the
    /// [`save()`](Workbook::save) method except that it returns the xlsx file
    /// as a `Vec<u8>` buffer suitable for streaming in a web application.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::SheetnameReused`] - Worksheet name is already in use in
    ///   the workbook.
    /// * [`XlsxError::IoError`] - A wrapper for various IO errors when creating
    ///   the xlsx file, or its sub-files.
    /// * [`XlsxError::ZipError`] - A wrapper for various zip errors when
    ///   creating the xlsx file, or its sub-files.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a simple workbook to a
    /// `Vec<u8>` buffer.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save_to_buffer.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     let buf = workbook.save_to_buffer()?;
    ///
    ///     println!("File size: {}", buf.len());
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn save_to_buffer(&mut self) -> Result<Vec<u8>, XlsxError> {
        let mut buf = vec![];
        let cursor = Cursor::new(&mut buf);
        self.save_internal(cursor)?;
        Ok(buf)
    }

    /// Save the Workbook as an xlsx file to a user supplied file/buffer.
    ///
    /// The workbook `save_to_writer()` method is similar to the
    /// [`save()`](Workbook::save) method except that it writes the xlsx file to
    /// types that implement the [`Write`] trait such as the [`std::fs::File`]
    /// type or buffers.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::SheetnameReused`] - Worksheet name is already in use in
    ///   the workbook.
    /// * [`XlsxError::IoError`] - A wrapper for various IO errors when creating
    ///   the xlsx file, or its sub-files.
    /// * [`XlsxError::ZipError`] - A wrapper for various zip errors when
    ///   creating the xlsx file, or its sub-files.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a simple workbook to some
    /// types that implement the `Write` trait like a file and a buffer.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save_to_writer.rs
    /// #
    /// # use std::fs::File;
    /// # use std::io::{Cursor, Write};
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     // Save the file to a File object.
    ///     let file = File::create("workbook1.xlsx")?;
    ///     workbook.save_to_writer(file)?;
    ///
    ///     // Save the file to a buffer. It is wrapped in a Cursor because it need to
    ///     // implement the `Seek` trait.
    ///     let mut cursor = Cursor::new(Vec::new());
    ///     workbook.save_to_writer(&mut cursor)?;
    ///
    ///     // Write the buffer to a file for the sake of the example.
    ///     let buf = cursor.into_inner();
    ///     let mut file = File::create("workbook2.xlsx")?;
    ///     Write::write_all(&mut file, &buf)?;
    ///
    ///     Ok(())
    /// }
    ///
    pub fn save_to_writer<W>(&mut self, writer: W) -> Result<(), XlsxError>
    where
        W: Write + Seek + Send,
    {
        self.save_internal(writer)?;
        Ok(())
    }

    /// Create a defined name in the workbook to use as a variable.
    ///
    /// The `define_name()` method is used to defined a variable name that can
    /// be used to represent a value, a single cell or a range of cells in a
    /// workbook. These are sometimes referred to as a "Named Ranges".
    ///
    /// Defined names are generally used to simplify or clarify formulas by
    /// using descriptive variable names. For example:
    ///
    /// ```text
    ///     // Global workbook name.
    ///     workbook.define_name("Exchange_rate", "=0.96")?;
    ///     worksheet.write_formula(0, 0, "=Exchange_rate")?;
    /// ```
    ///
    /// A name defined like this is "global" to the workbook and can be used in
    /// any worksheet in the workbook.  It is also possible to define a
    /// local/worksheet name by prefixing it with the sheet name using the
    /// syntax `"sheetname!defined_name"`:
    ///
    /// ```text
    ///     // Local worksheet name.
    ///     workbook.define_name('Sheet2!Sales', '=Sheet2!$G$1:$G$10')?;
    /// ```
    ///
    /// See the full example below.
    ///
    /// Note, Excel has limitations on names used in defined names. For example
    /// it must start with a letter or underscore and cannot contain a space or
    /// any of the characters: `,/*[]:\"'`. It also cannot look like an Excel
    /// range such as `A1`, `XFD12345` or `R1C1`. If in doubt it best to test
    /// the name in Excel first.
    ///
    /// For local defined names sheet name must exist (at the time of saving)
    /// and if the sheet name contains spaces or special characters you must
    /// follow the Excel convention and enclose it in single quotes:
    ///
    /// ```text
    ///     workbook.define_name("'New Data'!Sales", ""=Sheet2!$G$1:$G$10")?;
    /// ```
    ///
    /// The rules for names in Excel are explained in the Microsoft Office
    /// documentation on how to [Define and use names in
    /// formulas](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64)
    /// and subsections.
    ///
    /// # Parameters
    ///
    /// * `name` - The variable name to define.
    /// * `formula` - The formula, value or range that the name defines..
    ///
    /// # Errors
    ///
    /// * [`XlsxError::ParameterError`] - The following Excel error cases will
    ///   raise a `ParameterError` error:
    ///   * If the name doesn't start with a letter or underscore.
    ///   * If the name contains `,/*[]:\"'` or `space`.
    ///
    /// # Examples
    ///
    /// Example of how to create defined names using the `rust_xlsxwriter`
    /// library.
    ///
    /// This functionality is used to define user friendly variable names to
    /// represent a value, a single cell,  or a range of cells in a workbook.
    ///
    /// ```
    /// # // This code is available in examples/app_defined_name.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add two worksheets to the workbook.
    /// #     let _worksheet1 = workbook.add_worksheet();
    /// #     let _worksheet2 = workbook.add_worksheet();
    /// #
    ///     // Define some global/workbook names.
    ///     workbook.define_name("Exchange_rate", "=0.96")?;
    ///     workbook.define_name("Sales", "=Sheet1!$G$1:$H$10")?;
    ///
    ///     // Define a local/worksheet name. Over-rides the "Sales" name above.
    ///     workbook.define_name("Sheet2!Sales", "=Sheet2!$G$1:$G$10")?;
    ///
    /// #     // Write some text in the file and one of the defined names in a formula.
    /// #     for worksheet in workbook.worksheets_mut() {
    /// #         worksheet.set_column_width(0, 45)?;
    /// #         worksheet.write_string(0, 0, "This worksheet contains some defined names.")?;
    /// #         worksheet.write_string(1, 0, "See Formulas -> Name Manager above.")?;
    /// #         worksheet.write_string(2, 0, "Example formula in cell B3 ->")?;
    /// #
    /// #         worksheet.write_formula(2, 1, "=Exchange_rate")?;
    /// #     }
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("defined_name.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_defined_name1.png">
    ///
    /// Here is the output in the Excel Name Manager. Note that there is a
    /// Global/Workbook "Sales" variable name and a Local/Worksheet version.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_defined_name2.png">
    ///
    pub fn define_name(
        &mut self,
        name: impl Into<String>,
        formula: &str,
    ) -> Result<&mut Workbook, XlsxError> {
        let mut defined_name = DefinedName::new();
        let name = name.into();

        // Match Global/Workbook or Local/Worksheet defined names.
        match name.find('!') {
            Some(position) => {
                defined_name.quoted_sheet_name = name[0..position].to_string();
                defined_name.name = name[position + 1..].to_string();
                defined_name.name_type = DefinedNameType::Local;
            }
            None => {
                defined_name.name = name.to_string();
                defined_name.name_type = DefinedNameType::Global;
            }
        }

        // Excel requires that the name starts with a letter or underscore.
        // Also, backspace is allowed but undocumented by Excel.
        if !defined_name.name.chars().next().unwrap().is_alphabetic()
            && !defined_name.name.starts_with('_')
            && !defined_name.name.starts_with('\\')
        {
            let error = format!(
                "Name '{}' must start with a letter or underscore in Excel",
                defined_name.name
            );
            return Err(XlsxError::ParameterError(error));
        }

        // Excel also prohibits certain characters in the name.
        if defined_name
            .name
            .contains([' ', ',', '/', '*', '[', ']', ':', '"', '\''])
        {
            let error = format!(
                "Name '{}' cannot contain any of the characters `,/*[]:\"'` or `space` in Excel",
                defined_name.name
            );
            return Err(XlsxError::ParameterError(error));
        }

        defined_name.range = utility::formula_to_string(formula);
        defined_name.set_sort_name();

        self.user_defined_names.push(defined_name);

        Ok(self)
    }

    /// Set the Excel document metadata properties.
    ///
    /// Set various Excel document metadata properties such as Author or
    /// Creation Date. It is used in conjunction with the [`DocProperties`]
    /// struct.
    ///
    /// # Parameters
    ///
    /// * `properties` - A reference to a [`DocProperties`] object.
    ///
    /// # Examples
    ///
    /// An example of setting workbook document properties for a file created
    /// using the `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/app_doc_properties.rs
    /// #
    /// # use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let properties = DocProperties::new()
    ///         .set_title("This is an example spreadsheet")
    ///         .set_subject("That demonstrates document properties")
    ///         .set_author("A. Rust User")
    ///         .set_manager("J. Alfred Prufrock")
    ///         .set_company("Rust Solutions Inc")
    ///         .set_category("Sample spreadsheets")
    ///         .set_keywords("Sample, Example, Properties")
    ///         .set_comment("Created with Rust and rust_xlsxwriter");
    ///
    ///     workbook.set_properties(&properties);
    ///
    /// #     let worksheet = workbook.add_worksheet();
    ///
    /// #     worksheet.set_column_width(0, 30)?;
    /// #     worksheet.write_string(0, 0, "See File -> Info -> Properties")?;
    /// #
    /// #     workbook.save("doc_properties.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_doc_properties.png">
    ///
    ///
    /// The document properties can also be used to set a constant creation date
    /// so that a file generated by a `rust_xlsxwriter` program will have the
    /// same checksum no matter when it is created.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_properties_checksum2.rs
    /// #
    /// use rust_xlsxwriter::{DocProperties, ExcelDateTime, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Create a file creation date for the file.
    ///     let date = ExcelDateTime::from_ymd(2023, 1, 1)?;
    ///
    ///     // Add it to the document metadata.
    ///     let properties = DocProperties::new().set_creation_datetime(&date);
    ///     workbook.set_properties(&properties);
    ///
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///
    ///     workbook.save("properties.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    ///  See also [Checksum of a saved
    ///  file](../workbook/index.html#checksum-of-a-saved-file).
    ///
    pub fn set_properties(&mut self, properties: &DocProperties) -> &mut Workbook {
        self.properties = properties.clone();
        self
    }

    /// Add a recommendation to open the file in “read-only” mode.
    ///
    /// This method can be used to set the Excel “Read-only Recommended” option
    /// that is available when saving a file. This presents the user of the file
    /// with an option to open it in "read-only" mode. This means that any
    /// changes to the file can’t be saved back to the same file and must be
    /// saved to a new file.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a simple workbook which opens
    /// with a recommendation that the file should be opened in read only mode.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_read_only_recommended.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let _worksheet = workbook.add_worksheet();
    ///
    ///     workbook.read_only_recommended();
    ///
    ///     workbook.save("workbook.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Alert when you open the output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/workbook_read_only_recommended.png">
    ///
    pub fn read_only_recommended(&mut self) -> &mut Workbook {
        self.read_only_mode = 2;
        self
    }

    // -----------------------------------------------------------------------
    // Internal function/methods.
    // -----------------------------------------------------------------------

    // Reset workbook between saves.
    fn reset(&mut self) {
        self.writer.reset();

        self.xf_indices = HashMap::from([(Format::default(), 0)]);
        self.xf_formats = vec![Format::default()];
        self.dxf_indices = HashMap::new();
        self.dxf_formats = vec![];
        self.font_count = 0;
        self.fill_count = 0;
        self.border_count = 0;
        self.num_formats = vec![];

        for worksheet in &mut self.worksheets {
            worksheet.reset();
        }
    }

    // Internal function to prepare the workbook and other component files for
    // writing to the xlsx file.
    #[allow(clippy::similar_names)]
    fn save_internal<W: Write + Seek + Send>(&mut self, writer: W) -> Result<(), XlsxError> {
        // Reset workbook and worksheet state data between saves.
        self.reset();

        // Ensure that there is at least one worksheet in the workbook.
        if self.worksheets.is_empty() {
            self.add_worksheet();
        }
        // Ensure one sheet is active/selected.
        self.set_active_worksheets();

        // Check for the use of hyperlink style in the worksheets and if so add
        // a hyperlink style to the global formats.
        for worksheet in &self.worksheets {
            if worksheet.has_hyperlink_style {
                let format = Format::new().set_hyperlink();
                self.xf_indices.insert(format.clone(), 1);
                self.xf_formats.push(format);
                self.has_hyperlink_style = true;
                break;
            }
        }

        // Check for duplicate sheet names, which aren't allowed by Excel
        let mut unique_worksheet_names = HashSet::new();
        for worksheet in &self.worksheets {
            let worksheet_name = worksheet.name.to_lowercase();
            if unique_worksheet_names.contains(&worksheet_name) {
                return Err(XlsxError::SheetnameReused(worksheet_name));
            }

            unique_worksheet_names.insert(worksheet_name);
        }

        // Write any Tables associated with serialization areas.
        #[cfg(feature = "serde")]
        for worksheet in &mut self.worksheets {
            worksheet.store_serialized_tables()?;
        }

        // Convert any worksheet local formats to workbook/global formats. At
        // the worksheet level each unique format will have an index like 0, 1,
        // 2, etc., starting from 0 for each worksheet. However, at a workbook
        // level they may have an equivalent index of 1, 7, 5 or whatever
        // workbook order they appear in.
        let mut worksheet_xf_formats: Vec<Vec<Format>> = vec![];
        let mut worksheet_dxf_formats: Vec<Vec<Format>> = vec![];
        for worksheet in &self.worksheets {
            let formats = worksheet.xf_formats.clone();
            worksheet_xf_formats.push(formats);
            let formats = worksheet.dxf_formats.clone();
            worksheet_dxf_formats.push(formats);
        }

        let mut worksheet_xf_indices: Vec<Vec<u32>> = vec![];
        for formats in &worksheet_xf_formats {
            let mut indices = vec![];
            for format in formats {
                let index = self.format_xf_index(format);
                indices.push(index);
            }
            worksheet_xf_indices.push(indices);
        }
        let mut worksheet_dxf_indices: Vec<Vec<u32>> = vec![];
        for formats in &worksheet_dxf_formats {
            let mut indices = vec![];
            for format in formats {
                let index = self.format_dxf_index(format);
                indices.push(index);
            }
            worksheet_dxf_indices.push(indices);
        }

        for (i, worksheet) in self.worksheets.iter_mut().enumerate() {
            // Map worksheet/local format indices to the workbook/global values.
            worksheet.set_global_xf_indices(&worksheet_xf_indices[i]);
            worksheet.set_global_dxf_indices(&worksheet_dxf_indices[i]);

            // Perform the autofilter row hiding.
            worksheet.hide_autofilter_rows();

            // Set the index of the worksheets.
            worksheet.sheet_index = i;
        }

        // Generate a global array of embedded images from the worksheets.
        self.prepare_embedded_images();

        // Convert the images in the workbooks into drawing files and rel links.
        self.prepare_drawings();

        // Fill the chart data caches from worksheet data.
        self.prepare_chart_cache_data()?;

        // Prepare the formats for writing with styles.rs.
        self.prepare_format_properties();

        // Prepare worksheet tables.
        self.prepare_tables()?;

        // Collect workbook level metadata to help generate the xlsx file.
        let mut package_options = PackagerOptions::new();
        package_options = self.set_package_options(package_options)?;

        // Create the Packager object that will assemble the zip/xlsx file.
        let mut packager = Packager::new(writer);
        packager.assemble_file(self, &package_options)?;

        Ok(())
    }

    // Iterates through the worksheets and find which is the user defined Active
    // sheet. If none has been set then default to the first sheet, like Excel.
    fn set_active_worksheets(&mut self) {
        let mut active_index = 0;

        for (i, worksheet) in self.worksheets.iter().enumerate() {
            if worksheet.active {
                active_index = i;
            }
            if worksheet.first_sheet {
                self.first_sheet = i as u16;
            }
        }
        self.worksheets[active_index].set_active(true);
        self.active_tab = active_index as u16;
    }

    // Convert any embedded images in the worksheets to a global reference. Each
    // worksheet like have a local index to an embedded cell image. We need to
    // map these local references to a worksheet/global id that takes into
    // account duplicate images.
    fn prepare_embedded_images(&mut self) {
        let mut embedded_images = vec![];
        let mut image_ids: HashMap<u64, u32> = HashMap::new();
        let mut global_image_id = 0;

        for worksheet in &mut self.worksheets {
            if worksheet.embedded_images.is_empty() {
                continue;
            }

            let mut global_embedded_image_ids = vec![];
            for image in &worksheet.embedded_images {
                let image_id = match image_ids.get(&image.hash) {
                    Some(image_id) => *image_id,
                    None => {
                        global_image_id += 1;
                        embedded_images.push(image.clone());
                        image_ids.insert(image.hash, global_image_id);
                        global_image_id
                    }
                };

                global_embedded_image_ids.push(image_id);
            }

            worksheet.global_embedded_image_indices = global_embedded_image_ids;
        }

        self.embedded_images = embedded_images;
    }

    // Convert the images in the workbooks into drawing files and rel links.
    fn prepare_drawings(&mut self) {
        let mut chart_id = 1;
        let mut drawing_id = 1;
        let mut vml_drawing_id = 1;
        let mut image_id = self.embedded_images.len() as u32;

        // These are the image ids for each unique image file.
        let mut worksheet_image_ids: HashMap<u64, u32> = HashMap::new();
        let mut header_footer_image_ids: HashMap<u64, u32> = HashMap::new();

        for worksheet in &mut self.worksheets {
            if !worksheet.images.is_empty() {
                worksheet.prepare_worksheet_images(
                    &mut worksheet_image_ids,
                    &mut image_id,
                    drawing_id,
                );
            }

            if !worksheet.charts.is_empty() {
                chart_id = worksheet.prepare_worksheet_charts(chart_id, drawing_id);
            }

            // Increase the drawing number/id for image/chart file.
            if !worksheet.images.is_empty() || !worksheet.charts.is_empty() {
                drawing_id += 1;
            }

            if worksheet.has_header_footer_images() {
                // The header/footer images are counted from the last worksheet id.
                let base_image_id = worksheet_image_ids.len() as u32;

                worksheet.prepare_header_footer_images(
                    &mut header_footer_image_ids,
                    base_image_id,
                    vml_drawing_id,
                );
                vml_drawing_id += 1;
            }
        }
    }

    // Prepare and check each table in the workbook.
    fn prepare_tables(&mut self) -> Result<(), XlsxError> {
        let mut table_id = 1;
        let mut seen_table_names = HashSet::new();

        // Set a unique table id and table name and also set the .rel file
        // linkages.
        for worksheet in &mut self.worksheets {
            if !worksheet.tables.is_empty() {
                table_id = worksheet.prepare_worksheet_tables(table_id);
            }
        }

        // Check for duplicate table names.
        for worksheet in &self.worksheets {
            for table in &worksheet.tables {
                if seen_table_names.contains(&table.name.to_lowercase()) {
                    return Err(XlsxError::TableNameReused(table.name.to_string()));
                }

                seen_table_names.insert(table.name.to_lowercase());
            }
        }

        Ok(())
    }

    // Add worksheet number/string cache data to chart ranges. This isn't
    // strictly necessary but it helps non-Excel apps to render charts
    // correctly.
    fn prepare_chart_cache_data(&mut self) -> Result<(), XlsxError> {
        // First build up a hash of the chart data ranges. The data may not be
        // in the same worksheet as the chart so we need to do the lookup at the
        // workbook level.
        let mut chart_caches: HashMap<
            (String, RowNum, ColNum, RowNum, ColNum),
            ChartRangeCacheData,
        > = HashMap::new();

        // Add the chart ranges to the cache lookup table.
        for worksheet in &self.worksheets {
            if !worksheet.charts.is_empty() {
                for chart in worksheet.charts.values() {
                    Self::insert_chart_ranges_to_cache(chart, &mut chart_caches);

                    if let Some(chart) = &chart.combined_chart {
                        Self::insert_chart_ranges_to_cache(chart, &mut chart_caches);
                    }
                }
            }
        }

        // Populate the caches with data from the worksheet ranges.
        for (key, cache) in &mut chart_caches {
            if let Ok(worksheet) = self.worksheet_from_name(&key.0) {
                *cache = worksheet.get_cache_data(key.1, key.2, key.3, key.4);
            } else {
                let sheet_name = key.0.clone();
                let range = utility::chart_range_abs(&key.0, key.1, key.2, key.3, key.4);
                let error =
                    format!("Unknown worksheet name '{sheet_name}' in chart range '{range}'");

                return Err(XlsxError::UnknownWorksheetNameOrIndex(error));
            }
        }

        // Fill the caches back into the chart ranges.
        for worksheet in &mut self.worksheets {
            if !worksheet.charts.is_empty() {
                for chart in worksheet.charts.values_mut() {
                    Self::update_chart_ranges_from_cache(chart, &mut chart_caches);

                    if let Some(chart) = &mut chart.combined_chart {
                        Self::update_chart_ranges_from_cache(chart, &mut chart_caches);
                    }
                }
            }
        }

        Ok(())
    }

    // Insert all the various chart ranges into the lookup range cache.
    fn insert_chart_ranges_to_cache(
        chart: &Chart,
        chart_caches: &mut HashMap<(String, RowNum, ColNum, RowNum, ColNum), ChartRangeCacheData>,
    ) {
        Self::insert_to_chart_cache(&chart.title.range, chart_caches);
        Self::insert_to_chart_cache(&chart.x_axis.title.range, chart_caches);
        Self::insert_to_chart_cache(&chart.y_axis.title.range, chart_caches);

        for series in &chart.series {
            Self::insert_to_chart_cache(&series.title.range, chart_caches);
            Self::insert_to_chart_cache(&series.value_range, chart_caches);
            Self::insert_to_chart_cache(&series.category_range, chart_caches);

            for data_label in &series.custom_data_labels {
                Self::insert_to_chart_cache(&data_label.title.range, chart_caches);
            }

            if let Some(error_bars) = &series.y_error_bars {
                Self::insert_to_chart_cache(&error_bars.plus_range, chart_caches);
                Self::insert_to_chart_cache(&error_bars.minus_range, chart_caches);
            }

            if let Some(error_bars) = &series.x_error_bars {
                Self::insert_to_chart_cache(&error_bars.plus_range, chart_caches);
                Self::insert_to_chart_cache(&error_bars.minus_range, chart_caches);
            }
        }
    }

    // Update all the various chart ranges from the lookup range cache.
    fn update_chart_ranges_from_cache(
        chart: &mut Chart,
        chart_caches: &mut HashMap<(String, RowNum, ColNum, RowNum, ColNum), ChartRangeCacheData>,
    ) {
        Self::update_range_cache(&mut chart.title.range, chart_caches);
        Self::update_range_cache(&mut chart.x_axis.title.range, chart_caches);
        Self::update_range_cache(&mut chart.y_axis.title.range, chart_caches);

        for series in &mut chart.series {
            Self::update_range_cache(&mut series.title.range, chart_caches);
            Self::update_range_cache(&mut series.value_range, chart_caches);
            Self::update_range_cache(&mut series.category_range, chart_caches);

            for data_label in &mut series.custom_data_labels {
                if let Some(cache) = chart_caches.get(&data_label.title.range.key()) {
                    data_label.title.range.cache = cache.clone();
                }
            }

            if let Some(error_bars) = &mut series.y_error_bars {
                Self::update_range_cache(&mut error_bars.plus_range, chart_caches);
                Self::update_range_cache(&mut error_bars.minus_range, chart_caches);
            }

            if let Some(error_bars) = &mut series.x_error_bars {
                Self::update_range_cache(&mut error_bars.plus_range, chart_caches);
                Self::update_range_cache(&mut error_bars.minus_range, chart_caches);
            }
        }
    }

    // Insert a chart range (expressed as a hash/key value) into the chart cache
    // for lookup later.
    fn insert_to_chart_cache(
        range: &ChartRange,
        chart_caches: &mut HashMap<(String, RowNum, ColNum, RowNum, ColNum), ChartRangeCacheData>,
    ) {
        if range.has_data() {
            chart_caches.insert(range.key(), ChartRangeCacheData::new());
        }
    }

    // Populate a chart range cache with data read from the worksheet.
    fn update_range_cache(
        range: &mut ChartRange,
        chart_caches: &mut HashMap<(String, RowNum, ColNum, RowNum, ColNum), ChartRangeCacheData>,
    ) {
        if let Some(cache) = chart_caches.get(&range.key()) {
            range.cache = cache.clone();
        }
    }

    // Evaluate and clone formats from worksheets into a workbook level vector
    // of unique formats. Also return the index for use in remapping worksheet
    // format indices.
    fn format_xf_index(&mut self, format: &Format) -> u32 {
        match self.xf_indices.get_mut(format) {
            Some(xf_index) => *xf_index,
            None => {
                let xf_index = self.xf_formats.len() as u32;
                self.xf_formats.push(format.clone());
                self.xf_indices.insert(format.clone(), xf_index);
                xf_index
            }
        }
    }

    fn format_dxf_index(&mut self, format: &Format) -> u32 {
        match self.dxf_indices.get_mut(format) {
            Some(dxf_index) => *dxf_index,
            None => {
                let dxf_index = self.dxf_formats.len() as u32;
                self.dxf_formats.push(format.clone());
                self.dxf_indices.insert(format.clone(), dxf_index);
                dxf_index
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

    // Set the font index for the format objects. This only needs to be done for
    // XF formats. DXF formats are handled differently.
    fn prepare_fonts(&mut self) {
        let mut font_count: u16 = 0;
        let mut font_indices: HashMap<Font, u16> = HashMap::new();

        for xf_format in &mut self.xf_formats {
            match font_indices.get(&xf_format.font) {
                Some(font_index) => {
                    xf_format.set_font_index(*font_index, false);
                }
                None => {
                    font_indices.insert(xf_format.font.clone(), font_count);
                    xf_format.set_font_index(font_count, true);
                    font_count += 1;
                }
            }
        }
        self.font_count = font_count;
    }

    // Set the fill index for the format objects.
    fn prepare_fills(&mut self) {
        // The user defined fill properties start from 2 since there are 2
        // default fills: patternType="none" and patternType="gray125". The
        // following code adds these 2 default fills.
        let mut fill_count: u16 = 2;

        let mut fill_indices = HashMap::from([
            (Fill::default(), 0),
            (
                Fill {
                    pattern: crate::FormatPattern::Gray125,
                    ..Default::default()
                },
                1,
            ),
        ]);

        for xf_format in &mut self.xf_formats {
            let fill = &mut xf_format.fill;
            // For a solid fill (pattern == "solid") Excel reverses the role of
            // foreground and background colors, and
            if fill.pattern == FormatPattern::Solid
                && fill.background_color != Color::Default
                && fill.foreground_color != Color::Default
            {
                mem::swap(&mut fill.foreground_color, &mut fill.background_color);
            }

            // If the user specifies a foreground or background color without a
            // pattern they probably wanted a solid fill, so we fill in the
            // defaults.
            if (fill.pattern == FormatPattern::None || fill.pattern == FormatPattern::Solid)
                && fill.background_color != Color::Default
                && fill.foreground_color == Color::Default
            {
                fill.foreground_color = fill.background_color;
                fill.background_color = Color::Default;
                fill.pattern = FormatPattern::Solid;
            }

            if (fill.pattern == FormatPattern::None || fill.pattern == FormatPattern::Solid)
                && fill.background_color == Color::Default
                && fill.foreground_color != Color::Default
            {
                fill.background_color = Color::Default;
                fill.pattern = FormatPattern::Solid;
            }

            // Find unique or repeated fill ids.
            match fill_indices.get(fill) {
                Some(fill_index) => {
                    xf_format.set_fill_index(*fill_index, false);
                }
                None => {
                    fill_indices.insert(fill.clone(), fill_count);
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
        let mut border_indices: HashMap<Border, u16> = HashMap::new();

        for xf_format in &mut self.xf_formats {
            match border_indices.get(&xf_format.borders) {
                Some(border_index) => {
                    xf_format.set_border_index(*border_index, false);
                }
                None => {
                    border_indices.insert(xf_format.borders.clone(), border_count);
                    xf_format.set_border_index(border_count, true);
                    border_count += 1;
                }
            }
        }
        self.border_count = border_count;

        // For DXF borders we only need to check if any properties have changed.
        for dxf_format in &mut self.dxf_formats {
            dxf_format.has_border = !dxf_format.borders.is_default();
        }
    }

    // Set the number format index for the format objects.
    fn prepare_num_formats(&mut self) {
        let mut unique_num_formats: HashMap<String, u16> = HashMap::new();
        // User defined number formats in Excel start from index 164.
        let mut index = 164;
        let mut num_formats = vec![];
        let xf_formats = [&mut self.xf_formats, &mut self.dxf_formats];

        for xf_format in xf_formats.into_iter().flatten() {
            if xf_format.num_format_index > 0 {
                continue;
            }

            if xf_format.num_format.is_empty() {
                continue;
            }

            let num_format_string = xf_format.num_format.clone();

            match unique_num_formats.get(&num_format_string) {
                Some(index) => {
                    xf_format.set_num_format_index_u16(*index);
                }
                None => {
                    unique_num_formats.insert(num_format_string.clone(), index);
                    xf_format.set_num_format_index_u16(index);
                    index += 1;

                    // Only store XF formats (not DXF formats).
                    if !xf_format.is_dxf_format {
                        num_formats.push(num_format_string);
                    }
                }
            }
        }

        self.num_formats = num_formats;
    }

    // Collect some workbook level metadata to help generate the xlsx
    // package/file.
    fn set_package_options(
        &mut self,
        mut package_options: PackagerOptions,
    ) -> Result<PackagerOptions, XlsxError> {
        package_options.num_worksheets = self.worksheets.len() as u16;
        package_options.doc_security = self.read_only_mode;
        package_options.num_embedded_images = self.embedded_images.len() as u32;

        let mut defined_names = self.user_defined_names.clone();
        let mut sheet_names: HashMap<String, u16> = HashMap::new();

        // Iterate over the worksheets to capture workbook and update the
        // package options metadata.
        for (sheet_index, worksheet) in self.worksheets.iter().enumerate() {
            let sheet_name = worksheet.name.clone();
            let quoted_sheet_name = utility::quote_sheetname(&sheet_name);
            sheet_names.insert(sheet_name.clone(), sheet_index as u16);

            if worksheet.visible == Visible::VeryHidden {
                package_options.worksheet_names.push(String::new());
            } else {
                package_options.worksheet_names.push(sheet_name.clone());
            }

            package_options.properties = self.properties.clone();

            if worksheet.uses_string_table {
                package_options.has_sst_table = true;
            }

            if worksheet.has_dynamic_arrays {
                package_options.has_metadata = true;
                package_options.has_dynamic_functions = true;
            }

            if !worksheet.embedded_images.is_empty() {
                package_options.has_metadata = true;
                package_options.has_embedded_images = true;
                if worksheet.has_embedded_image_descriptions {
                    package_options.has_embedded_image_descriptions = true;
                }
            }

            if worksheet.has_header_footer_images() {
                package_options.has_vml = true;
            }

            if !worksheet.drawing.drawings.is_empty() {
                package_options.num_drawings += 1;
            }

            if !worksheet.charts.is_empty() {
                package_options.num_charts += worksheet.charts.len() as u16;
            }

            if !worksheet.tables.is_empty() {
                package_options.num_tables += worksheet.tables.len() as u16;
            }

            // Store the autofilter areas which are a category of defined name.
            if worksheet.autofilter_defined_name.in_use {
                let mut defined_name = worksheet.autofilter_defined_name.clone();
                defined_name.initialize(&quoted_sheet_name);
                defined_names.push(defined_name);
            }

            // Store any user defined print areas which are a category of defined name.
            if worksheet.print_area_defined_name.in_use {
                let mut defined_name = worksheet.print_area_defined_name.clone();
                defined_name.initialize(&quoted_sheet_name);
                defined_names.push(defined_name);
            }

            // Store any user defined print repeat rows/columns which are a
            // category of defined name.
            if worksheet.repeat_row_cols_defined_name.in_use {
                let mut defined_name = worksheet.repeat_row_cols_defined_name.clone();
                defined_name.initialize(&quoted_sheet_name);
                defined_names.push(defined_name);
            }

            // Set the used image types.
            for i in 0..NUM_IMAGE_FORMATS {
                if worksheet.image_types[i] {
                    package_options.image_types[i] = true;
                }
            }
        }

        // Map the sheet name and associated index so that we can map a sheet
        // reference in a Local/Sheet defined name to a worksheet index.
        for defined_name in &mut defined_names {
            let sheet_name = defined_name.unquoted_sheet_name();

            if !sheet_name.is_empty() {
                match sheet_names.get(&sheet_name) {
                    Some(index) => defined_name.index = *index,
                    None => {
                        let error = format!(
                            "Unknown worksheet name '{}' in defined name '{}'",
                            sheet_name, defined_name.name
                        );
                        return Err(XlsxError::ParameterError(error));
                    }
                }
            }
        }

        // Excel stores defined names in a sorted order.
        defined_names.sort_by_key(|n| (n.sort_name.clone(), n.range.clone()));

        // Map the non-Global defined names to App.xml entries.
        for defined_name in &defined_names {
            let app_name = defined_name.app_name();
            if !app_name.is_empty() {
                package_options.defined_names.push(app_name);
            }
        }

        self.defined_names = defined_names;

        Ok(package_options)
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the workbook element.
        self.write_workbook();

        // Write the fileVersion element.
        self.write_file_version();

        // Write the fileSharing element.
        if self.read_only_mode == 2 {
            self.write_file_sharing();
        }

        // Write the workbookPr element.
        self.write_workbook_pr();

        // Write the bookViews element.
        self.write_book_views();

        // Write the sheets element.
        self.write_sheets();

        // Write the definedNames element.
        if !self.defined_names.is_empty() {
            self.write_defined_names();
        }

        // Write the calcPr element.
        self.write_calc_pr();

        // Close the workbook tag.
        self.writer.xml_end_tag("workbook");
    }

    // Write the <workbook> element.
    fn write_workbook(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let xmlns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        let attributes = [("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag("workbook", &attributes);
    }

    // Write the <fileVersion> element.
    fn write_file_version(&mut self) {
        let attributes = [
            ("appName", "xl"),
            ("lastEdited", "4"),
            ("lowestEdited", "4"),
            ("rupBuild", "4505"),
        ];

        self.writer.xml_empty_tag("fileVersion", &attributes);
    }

    // Write the <fileSharing> element.
    fn write_file_sharing(&mut self) {
        let attributes = [("readOnlyRecommended", "1")];

        self.writer.xml_empty_tag("fileSharing", &attributes);
    }

    // Write the <workbookPr> element.
    fn write_workbook_pr(&mut self) {
        let attributes = [("defaultThemeVersion", "124226")];

        self.writer.xml_empty_tag("workbookPr", &attributes);
    }

    // Write the <bookViews> element.
    fn write_book_views(&mut self) {
        self.writer.xml_start_tag_only("bookViews");

        // Write the workbookView element.
        self.write_workbook_view();

        self.writer.xml_end_tag("bookViews");
    }

    // Write the <workbookView> element.
    fn write_workbook_view(&mut self) {
        let mut attributes = vec![
            ("xWindow", "240".to_string()),
            ("yWindow", "15".to_string()),
            ("windowWidth", "16095".to_string()),
            ("windowHeight", "9660".to_string()),
        ];

        // Store the firstSheet attribute when it isn't the first sheet.
        if self.first_sheet > 0 {
            let first_sheet = self.first_sheet + 1;
            attributes.push(("firstSheet", first_sheet.to_string()));
        }

        // Store the activeTab attribute when it isn't the first sheet.
        if self.active_tab > 0 {
            attributes.push(("activeTab", self.active_tab.to_string()));
        }

        self.writer.xml_empty_tag("workbookView", &attributes);
    }

    // Write the <sheets> element.
    fn write_sheets(&mut self) {
        self.writer.xml_start_tag_only("sheets");

        let mut worksheet_data = vec![];
        for worksheet in &self.worksheets {
            worksheet_data.push((worksheet.name.clone(), worksheet.visible));
        }

        for (index, data) in worksheet_data.iter().enumerate() {
            // Write the sheet element.
            self.write_sheet(&data.0, data.1, (index + 1) as u16);
        }

        self.writer.xml_end_tag("sheets");
    }

    // Write the <sheet> element.
    fn write_sheet(&mut self, name: &str, visible: Visible, index: u16) {
        let sheet_id = format!("{index}");
        let ref_id = format!("rId{index}");

        let mut attributes = vec![("name", name.to_string()), ("sheetId", sheet_id)];

        match visible {
            Visible::Default => {}
            Visible::Hidden => attributes.push(("state", "hidden".to_string())),
            Visible::VeryHidden => attributes.push(("state", "veryHidden".to_string())),
        }

        attributes.push(("r:id", ref_id));

        self.writer.xml_empty_tag("sheet", &attributes);
    }

    // Write the <definedNames> element.
    fn write_defined_names(&mut self) {
        self.writer.xml_start_tag_only("definedNames");

        for defined_name in &self.defined_names {
            let mut attributes = vec![("name", defined_name.name())];

            match defined_name.name_type {
                DefinedNameType::Global => {}
                _ => {
                    attributes.push(("localSheetId", defined_name.index.to_string()));
                }
            }

            if let DefinedNameType::Autofilter = defined_name.name_type {
                attributes.push(("hidden", "1".to_string()));
            }

            self.writer
                .xml_data_element("definedName", &defined_name.range, &attributes);
        }

        self.writer.xml_end_tag("definedNames");
    }

    // Write the <calcPr> element.
    fn write_calc_pr(&mut self) {
        let attributes = [("calcId", "124519"), ("fullCalcOnLoad", "1")];

        self.writer.xml_empty_tag("calcPr", &attributes);
    }
}
