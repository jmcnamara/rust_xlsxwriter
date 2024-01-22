// Entry point for `rust_xlsxwriter` library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![doc(html_logo_url = "https://rustxlsxwriter.github.io/images/rust_xlsxwriter_logo.png")]
#![cfg_attr(docsrs, feature(doc_cfg))]

//! `rust_xlsxwriter` is a Rust library for writing Excel files in the xlsx
//! format.
//!
//! <img src="https://rustxlsxwriter.github.io/images/demo.png">
//!
//! The `rust_xlsxwriter` crate can be used to write text, numbers, dates and
//! formulas to multiple worksheets in a new Excel 2007+ xlsx file. It has a
//! focus on performance and on fidelity with the file format created by Excel.
//! It cannot be used to modify an existing file.
//!
//! `rust_xlsxwriter` is a port of the [`XlsxWriter`] Python module by the same
//! author. Feature porting is a work in progress. The currently supported
//! features are:
//!
//! - Support for writing all basic Excel data types.
//! - Full cell formatting support.
//! - Formula support, including new Excel 365 dynamic functions.
//! - Charts.
//! - Hyperlink support.
//! - Page/Printing Setup support.
//! - Merged ranges.
//! - Conditional formatting.
//! - Worksheet PNG/JPEG/GIF/BMP images.
//! - Rich multi-format strings.
//! - Defined names.
//! - Autofilters.
//! - Worksheet Tables.
//!
//! [`XlsxWriter`]: https://xlsxwriter.readthedocs.io/index.html
//!
//! # Table of contents
//!
//! - [`Tutorial`](crate::tutorial): A getting started and tutorial guide.
//! - [`Cookbook`](crate::cookbook): Examples of using `rust_xlsxwriter`.
//! - [`Workbook`]: The entry point for creating an Excel workbook of
//!   worksheets.
//! - [`Worksheet`]: The main spreadsheet canvas for writing data and objects to
//!   a worksheet.
//! - [`Format`]: The interface for adding formatting to worksheets and other
//!   objects.
//! - [`Chart`] struct: The interface for creating worksheet charts.
//!   - [`Working with charts`](crate::chart).
//! - [`Table`]: The interface for worksheet tables.
//! - [`Image`]: The interface for images used in worksheets.
//! - [`Conditional Formats`](crate::conditional_format): Working with
//!   conditional formatting in worksheets.
//! - [`ExcelDateTime`]: A type to represent dates and times in Excel format.
//! - [`Formula`]: A type for Excel formulas.
//! - [`Url`]: A type for URLs/Hyperlinks used in worksheets.
//! - [`DocProperties`]: The interface used to create an object to represent
//!   document metadata properties.
//!
//! Other external documentation:
//!
//! - [User Guide]: Working with the `rust_xlsxwriter` library.
//! - [Release Notes].
//! - [Roadmap of planned features].
//!
//! [User Guide]: https://rustxlsxwriter.github.io/index.html
//! [Release Notes]: https://rustxlsxwriter.github.io/changelog.html
//! [Roadmap of planned features]:
//!     https://github.com/jmcnamara/rust_xlsxwriter/issues/1
//!
//! # Example
//!
//! <img src="https://rustxlsxwriter.github.io/images/demo.png">
//!
//! Sample code to generate the Excel file shown above.
//!
//! ```rust
//! # // This code is available in examples/app_demo.rs
//! #
//! use rust_xlsxwriter::*;
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Create some formats to use in the worksheet.
//!     let bold_format = Format::new().set_bold();
//!     let decimal_format = Format::new().set_num_format("0.000");
//!     let date_format = Format::new().set_num_format("yyyy-mm-dd");
//!     let merge_format = Format::new()
//!         .set_border(FormatBorder::Thin)
//!         .set_align(FormatAlign::Center);
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Set the column width for clarity.
//!     worksheet.set_column_width(0, 22)?;
//!
//!     // Write a string without formatting.
//!     worksheet.write(0, 0, "Hello")?;
//!
//!     // Write a string with the bold format defined above.
//!     worksheet.write_with_format(1, 0, "World", &bold_format)?;
//!
//!     // Write some numbers.
//!     worksheet.write(2, 0, 1)?;
//!     worksheet.write(3, 0, 2.34)?;
//!
//!     // Write a number with formatting.
//!     worksheet.write_with_format(4, 0, 3.00, &decimal_format)?;
//!
//!     // Write a formula.
//!     worksheet.write(5, 0, Formula::new("=SIN(PI()/4)"))?;
//!
//!     // Write a date.
//!     let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
//!     worksheet.write_with_format(6, 0, &date, &date_format)?;
//!
//!     // Write some links.
//!     worksheet.write(7, 0, Url::new("https://www.rust-lang.org"))?;
//!     worksheet.write(8, 0, Url::new("https://www.rust-lang.org").set_text("Rust"))?;
//!
//!     // Write some merged cells.
//!     worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;
//!
//!     // Insert an image.
//!     let image = Image::new("examples/rust_logo.png")?;
//!     worksheet.insert_image(1, 2, &image)?;
//!
//!     // Save the file to disk.
//!     workbook.save("demo.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! ## Crate Features
//!
//! The following is a list of the features supports by the `rust_xlsxwriter`
//! crate:
//!
//! - `default`: Includes all the standard functionality. Has dependencies on
//!   `zip`, `regex` and `lazy_static`.
//! - `serde`: Adds supports for Serde serialization. This is off by default.
//! - `chrono`: Adds supports for Chrono date/time types to the API. This is off
//!   by default.
//! - `zlib`: Adds a dependency on zlib and a C compiler. This includes the same
//!   features as `default` but is 1.5x faster for large files.
//! - `polars`: Add support for mapping between `PolarsError` and
//!   `rust_xlsxwriter::XlsxError` to make code that handles both types of error
//!   easier to write.
//! - `wasm`: Adds a dependency on `js-sys` and `wasm-bindgen` to allow
//!   compilation for wasm/JavaScript targets.
//!
//!
mod app;
mod content_types;
mod core;
mod custom;
mod datetime;
mod drawing;
mod error;
mod filter;
mod format;
mod formula;
mod image;
mod metadata;
mod packager;
mod properties;
mod protection;
mod relationship;
mod shared_strings;
mod shared_strings_table;
mod styles;
mod table;
mod theme;
mod url;
mod vml;
mod workbook;
mod worksheet;
mod xmlwriter;

#[cfg(feature = "serde")]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub mod serializer;

pub mod chart;
pub mod conditional_format;
pub mod cookbook;
pub mod tutorial;
pub mod utility;

#[cfg(test)]
mod test_functions;

// Re-export the public APIs.
pub use datetime::*;
pub use error::*;
pub use filter::*;
pub use format::*;
pub use formula::*;
pub use image::*;
pub use properties::*;
pub use protection::*;
pub use table::*;
pub use url::*;
pub use workbook::*;
pub use worksheet::*;

#[cfg(feature = "serde")]
#[doc(hidden)]
pub use serializer::*;

#[doc(hidden)]
pub use chart::*;
#[doc(hidden)]
pub use conditional_format::*;
#[doc(hidden)]
pub use utility::*;

#[macro_use]
extern crate lazy_static;

#[cfg(feature = "serde")]
extern crate rust_xlsxwriter_derive;

#[cfg(feature = "serde")]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub use rust_xlsxwriter_derive::XlsxSerialize;
