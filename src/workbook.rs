// workbook - A module for creating the Excel workbook.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::collections::HashMap;
use std::mem;
use std::path::Path;

use crate::error::XlsxError;
use crate::format::Format;
use crate::packager::Packager;
use crate::packager::PackagerOptions;
use crate::utility;
use crate::worksheet::Worksheet;
use crate::xmlwriter::XMLWriter;
use crate::{XlsxColor, XlsxPattern};

/// The workbook struct represents an Excel file in it's entirety. It is the
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
/// use rust_xlsxwriter::{Format, Workbook, XlsxAlign, XlsxBorder, XlsxError};
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
///         .set_border(XlsxBorder::Thin)
///         .set_align(XlsxAlign::Center);
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Set the column width for clarity.
///     worksheet.set_column_width(0, 22)?;
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
///     // Write a date.
///     let date = NaiveDate::from_ymd(2023, 1, 25);
///     worksheet.write_date(6, 0, date, &date_format)?;
///
///     // Write some links.
///     worksheet.write_url(7, 0, "https://www.rust-lang.org")?;
///     worksheet.write_url_with_text(8, 0, "https://www.rust-lang.org", "Learn Rust!")?;
///
///     // Write some merged cells.
///     worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;
///
///     // Save the file to disk.
///     workbook.save("demo.xlsx")?;
///
///     Ok(())
/// }
/// ```
pub struct Workbook {
    pub(crate) writer: XMLWriter,
    pub(crate) worksheets: Vec<Worksheet>,
    xf_indices: HashMap<String, u32>,
    pub(crate) xf_formats: Vec<Format>,
    pub(crate) font_count: u16,
    pub(crate) fill_count: u16,
    pub(crate) border_count: u16,
    pub(crate) num_format_count: u16,
    pub(crate) has_hyperlink_style: bool,
    active_tab: u16,
    first_sheet: u16,
    defined_names: Vec<DefinedName>,
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
    /// everything to an xlsx file with [`save()`](Workbook::save),
    /// [`save_to_path()`](Workbook::save_to_path) or
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
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     _ = workbook.add_worksheet();
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
        let default_format = Format::new();
        let xf_indices = HashMap::from([(default_format.format_key(), 0)]);

        Workbook {
            writer,
            worksheets: vec![],
            xf_indices,
            xf_formats: vec![default_format],
            font_count: 0,
            fill_count: 0,
            border_count: 0,
            num_format_count: 0,
            active_tab: 0,
            first_sheet: 0,
            defined_names: vec![],
            has_hyperlink_style: false,
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
    /// The `add_worksheet()` method returns a borrowed mutable reference to a
    /// Worksheet instance owned by the Workbook so only one worksheet can be in
    /// existence at a time, see the example below. This limitation can be
    /// avoided, if necessary, by creating standalone Worksheet objects via
    /// [`Worksheet::new()`] and then later adding them to the workbook with
    /// [`workbook.push_worksheet`](Workbook::push_worksheet).
    ///
    /// See also the `rust_xlsxwriter` documentation on [Creating worksheets]
    /// and working with the borrow checker.
    ///
    /// [Creating worksheets]: https://rustxlsxwriter.github.io/worksheet/create.html
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
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
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
    /// the `rust_xlsxwriter` documentation on [Creating worksheets]
    /// and working with the borrow checker.
    ///
    /// [Creating worksheets]: https://rustxlsxwriter.github.io/worksheet/create.html
    ///
    /// # Arguments
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
    ///     worksheet1.write_string_only(0, 0, "Hello")?;
    ///
    ///     // If we don't try to use the workbook1 reference again we can switch to
    ///     // using a reference to worksheet2.
    ///     let mut worksheet2 = workbook.add_worksheet();
    ///     worksheet2.write_string_only(0, 0, "Hello")?;
    ///
    ///     // Stop using worksheet2 and move back to worksheet1.
    ///     worksheet1 = workbook.worksheet_from_index(0)?;
    ///     worksheet1.write_string_only(1, 0, "Sheet1")?;
    ///
    ///     // Stop using worksheet1 and move back to worksheet2.
    ///     worksheet2 = workbook.worksheet_from_index(1)?;
    ///     worksheet2.write_string_only(1, 0, "Sheet2")?;
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
    /// the `rust_xlsxwriter` documentation on [Creating worksheets] and working
    /// with the borrow checker.
    ///
    /// # Arguments
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
    ///
    ///     // Start with a reference to worksheet1.
    ///     let mut worksheet1 = workbook.add_worksheet();
    ///     worksheet1.write_string_only(0, 0, "Hello")?;
    ///
    ///     // If we don't try to use the workbook1 reference again we can switch to
    ///     // using a reference to worksheet2.
    ///     let mut worksheet2 = workbook.add_worksheet();
    ///     worksheet2.write_string_only(0, 0, "Hello")?;
    ///
    ///     // Stop using worksheet2 and move back to worksheet1.
    ///     worksheet1 = workbook.worksheet_from_index(0)?;
    ///     worksheet1.write_string_only(1, 0, "Sheet1")?;
    ///
    ///     // Stop using worksheet1 and move back to worksheet2.
    ///     worksheet2 = workbook.worksheet_from_index(1)?;
    ///     worksheet2.write_string_only(1, 0, "Sheet2")?;
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
    /// See also the `rust_xlsxwriter` documentation on [Creating worksheets]
    /// and working with the borrow checker.
    ///
    /// [Creating worksheets]:
    ///     https://rustxlsxwriter.github.io/worksheet/create.html
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
    ///         worksheet.write_string_only(0, 0, "Hello")?;
    ///         worksheet.write_number_only(1, 0, 12345)?;
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
    ///
    ///     // Add three worksheets to the workbook.
    ///     let _ = workbook.add_worksheet();
    ///     let _ = workbook.add_worksheet();
    ///     let _ = workbook.add_worksheet();
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
    /// There are two way of creating a worksheet object with rust_xlsxwriter:
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
    /// See also the `rust_xlsxwriter` documentation on [Creating worksheets]
    /// and working with the borrow checker.
    ///
    /// [Creating worksheets]: https://rustxlsxwriter.github.io/worksheet/create.html
    ///
    /// # Arguments
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
    ///     worksheet.write_string_only(0, 0, "Hello")?;
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
    /// same Workbook object to different   file names. However, `save()` is an
    /// expensive operation which assembles multiple files into an xlsx/zip
    /// container so for performance reasons you shouldn't call it
    /// unnecessarily.
    ///
    /// # Arguments
    ///
    /// * `filename` - The name of the new Excel file to create.
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
    /// The following example demonstrates creating a simple workbook, with one
    /// unused worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     _ = workbook.add_worksheet();
    ///
    ///     workbook.save("workbook.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn save(&mut self, filename: &str) -> Result<(), XlsxError> {
        let filehandle = FileHandleFrom::String(filename);
        _ = self.save_internal(filehandle)?;
        Ok(())
    }

    /// Save the Workbook as an xlsx file using a Path reference.
    ///
    /// The workbook `save_to_path()` method writes all Workbook data to a new
    /// xlsx file using a a [`std::path`] Path or PathBuf instance. It will
    /// overwrite any existing file.
    ///
    /// For most cases the [`save()`](Workbook::save) method which uses a simple
    /// string representation of the file path/name will be sufficient. However,
    /// there are use cases, on Windows in particular, where generating the path
    /// string may be error prone and where it can be preferable to use a
    /// [`std::path`] Path or PathBuf instance and `save_to_path()`.
    ///
    /// # Arguments
    ///
    /// * `path` - A reference to a [`std::path`] Path or PathBuf instance.
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
    /// The following example demonstrates creating a simple workbook using a
    /// rust Path reference.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save_to_path.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let path = std::path::Path::new("workbook.xlsx");
    ///     let mut workbook = Workbook::new();
    ///
    ///     _ = workbook.add_worksheet();
    ///
    ///     workbook.save_to_path(&path)?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn save_to_path(&mut self, path: &Path) -> Result<(), XlsxError> {
        let filehandle = FileHandleFrom::Path(path);
        _ = self.save_internal(filehandle)?;
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
    /// The following example demonstrates creating a simple workbook to a Vec<u8>
    /// buffer.
    ///
    /// ```
    /// # // This code is available in examples/doc_workbook_save_to_buffer.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.write_string_only(0, 0, "Hello")?;
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
        let filehandle = FileHandleFrom::Buffer;
        let buf = self.save_internal(filehandle)?;
        Ok(buf)
    }

    // Set the index for the format. This is currently only used in testing but
    // may be used publicly at a later stage.
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

    // Internal function to prepare the workbook and other component files for
    // writing to the xlsx file.
    fn save_internal(&mut self, filehandle: FileHandleFrom) -> Result<Vec<u8>, XlsxError> {
        // Reset workbook and worksheet xml writers between saves.
        self.writer.reset();
        for worksheet in self.worksheets.iter_mut() {
            worksheet.writer.reset();
        }

        // Clear any global metadata arrays between saves.
        self.defined_names = vec![];

        // Ensure that there is at least one worksheet in the workbook.
        if self.worksheets.is_empty() {
            self.add_worksheet();
        }
        // Ensure one sheet is active/selected.
        self.set_active_worksheets();

        // Check for the use of hyperlink style in the worksheets and if so add
        // a hyperlink style to the global formats.
        for worksheet in self.worksheets.iter() {
            if worksheet.has_hyperlink_style {
                let format = Format::new().set_hyperlink();
                self.xf_indices.insert(format.format_key(), 1);
                self.xf_formats.push(format);
                self.has_hyperlink_style = true;
                break;
            }
        }

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

        // Collect workbook level metadata to help generate the xlsx file.
        let mut package_options = PackagerOptions::new();
        package_options = self.set_package_options(package_options)?;

        // Create the Packager object that will assemble the zip/xlsx file.
        let mut buf: Vec<u8> = vec![];
        match filehandle {
            FileHandleFrom::String(filename) => {
                let path = std::path::Path::new(filename);
                let file = std::fs::File::create(path)?;
                let mut packager = Packager::new(file)?;
                packager.assemble_file(self, &package_options)?;
            }
            FileHandleFrom::Path(path) => {
                let file = std::fs::File::create(path)?;
                let mut packager = Packager::new(file)?;
                packager.assemble_file(self, &package_options)?;
            }
            FileHandleFrom::Buffer => {
                let cursor = std::io::Cursor::new(&mut buf);
                let mut packager = Packager::new(cursor)?;
                packager.assemble_file(self, &package_options)?;
            }
        };

        Ok(buf)
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

    // Evaluate and clone formats from worksheets into a workbook level vector
    // of unique formats. Also return the index for use in remapping worksheet
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
                && xf_format.background_color.is_not_default()
                && xf_format.foreground_color.is_not_default()
            {
                mem::swap(
                    &mut xf_format.foreground_color,
                    &mut xf_format.background_color,
                );
            }

            // If the user specifies a foreground or background color without a
            // pattern they probably wanted a solid fill, so we fill in the
            // defaults.
            if (xf_format.pattern == XlsxPattern::None || xf_format.pattern == XlsxPattern::Solid)
                && xf_format.background_color.is_not_default()
                && xf_format.foreground_color.is_default()
            {
                xf_format.foreground_color = xf_format.background_color;
                xf_format.background_color = XlsxColor::Automatic;
                xf_format.pattern = XlsxPattern::Solid;
            }

            if (xf_format.pattern == XlsxPattern::None || xf_format.pattern == XlsxPattern::Solid)
                && xf_format.background_color.is_default()
                && xf_format.foreground_color.is_not_default()
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

    // Collect some workbook level metadata to help generate the xlsx
    // package/file.
    fn set_package_options(
        &mut self,
        mut package_options: PackagerOptions,
    ) -> Result<PackagerOptions, XlsxError> {
        package_options.num_worksheets = self.worksheets.len() as u16;

        // Iterate over the worksheets to capture workbook and update the
        // package options metadata.
        for (sheet_index, worksheet) in self.worksheets.iter().enumerate() {
            let sheet_name = worksheet.name.clone();
            let quoted_sheet_name = utility::quote_sheetname(&sheet_name);

            // Check for duplicate sheet names, which aren't allowed by Excel.
            if package_options.worksheet_names.contains(&sheet_name) {
                return Err(XlsxError::SheetnameReused(sheet_name));
            }

            package_options.worksheet_names.push(sheet_name);

            if worksheet.uses_string_table {
                package_options.has_sst_table = true;
            }

            if worksheet.has_dynamic_arrays {
                package_options.has_dynamic_arrays = true;
            }

            // Store any user defined print areas which are a category of defined name.
            if !worksheet.print_area.is_empty() {
                let defined_name = DefinedName {
                    name: "_xlnm.Print_Area".to_string(),
                    range: format!("{}!{}", quoted_sheet_name, worksheet.print_area),
                    index: sheet_index as u16,
                };

                self.defined_names.push(defined_name);
                package_options
                    .defined_names
                    .push(format!("{}!Print_Area", quoted_sheet_name));
            }

            // Store any user defined print repeat rows/columns which are a
            // category of defined name.
            if !worksheet.repeat_row_range.is_empty() || !worksheet.repeat_col_range.is_empty() {
                let range;

                if !worksheet.repeat_row_range.is_empty() && !worksheet.repeat_col_range.is_empty()
                {
                    range = format!(
                        "{}!{},{}!{}",
                        quoted_sheet_name,
                        worksheet.repeat_col_range,
                        quoted_sheet_name,
                        worksheet.repeat_row_range
                    );
                } else if !worksheet.repeat_row_range.is_empty() {
                    range = format!("{}!{}", quoted_sheet_name, worksheet.repeat_row_range);
                } else {
                    range = format!("{}!{}", quoted_sheet_name, worksheet.repeat_col_range);
                }

                let defined_name = DefinedName {
                    name: "_xlnm.Print_Titles".to_string(),
                    range: range.clone(),
                    index: sheet_index as u16,
                };

                self.defined_names.push(defined_name);
                package_options
                    .defined_names
                    .push(format!("{}!Print_Titles", quoted_sheet_name));
            }
        }

        Ok(package_options)
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

        self.writer.xml_empty_tag_attr("workbookView", &attributes);
    }

    // Write the <sheets> element.
    fn write_sheets(&mut self) {
        self.writer.xml_start_tag("sheets");

        let mut worksheet_data = vec![];
        for worksheet in self.worksheets.iter() {
            worksheet_data.push((worksheet.name.clone(), worksheet.hidden));
        }

        for (index, data) in worksheet_data.iter().enumerate() {
            // Write the sheet element.
            self.write_sheet(&data.0, data.1, (index + 1) as u16);
        }

        self.writer.xml_end_tag("sheets");
    }

    // Write the <sheet> element.
    fn write_sheet(&mut self, name: &str, is_hidden: bool, index: u16) {
        let sheet_id = format!("{}", index);
        let ref_id = format!("rId{}", index);

        let mut attributes = vec![("name", name.to_string()), ("sheetId", sheet_id)];

        if is_hidden {
            attributes.push(("state", "hidden".to_string()));
        }

        attributes.push(("r:id", ref_id));

        self.writer.xml_empty_tag_attr("sheet", &attributes);
    }

    // Write the <definedNames> element.
    fn write_defined_names(&mut self) {
        self.writer.xml_start_tag("definedNames");

        for defined_name in self.defined_names.iter() {
            let attributes = vec![
                ("name", defined_name.name.to_string()),
                ("localSheetId", defined_name.index.to_string()),
            ];

            self.writer
                .xml_data_element_attr("definedName", &defined_name.range, &attributes);
        }

        self.writer.xml_end_tag("definedNames");
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

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------

#[derive(Clone)]
struct DefinedName {
    name: String,
    range: String,
    index: u16,
}

pub(crate) enum FileHandleFrom<'a> {
    String(&'a str),
    Path(&'a Path),
    Buffer,
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::Workbook;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut workbook = Workbook::default();
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
