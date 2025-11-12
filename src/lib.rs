// Entry point for the `rust_xlsxwriter` library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#![doc(html_logo_url = "https://rustxlsxwriter.github.io/images/rust_xlsxwriter_logo.png")]
#![cfg_attr(docsrs, feature(doc_cfg))]

//! `rust_xlsxwriter` is a Rust library for writing Excel files in the xlsx
//! format.
//!
//! <img src="https://rustxlsxwriter.github.io/images/demo.png">
//!
//! The `rust_xlsxwriter` crate can be used to write text, numbers, dates, and
//! formulas to multiple worksheets in a new Excel 2007+ `.xlsx` file. It has a
//! focus on performance and fidelity with the file format created by Excel. It
//! cannot be used to modify an existing file.
//!
//! `rust_xlsxwriter` is a rewrite of the Python [`XlsxWriter`] library in Rust
//! by the same author, with additional Rust-like features and APIs. The
//! supported features are:
//!
//! - Support for writing all basic Excel data types.
//! - Full cell formatting support.
//! - Formula support, including new Excel 365 dynamic functions.
//! - Charts.
//! - Hyperlink support.
//! - Page/Printing Setup support.
//! - Merged ranges.
//! - Conditional formatting.
//! - Data validation.
//! - Cell Notes.
//! - Textboxes.
//! - Checkboxes.
//! - Sparklines.
//! - Worksheet PNG/JPEG/GIF/BMP images.
//! - Workbook themes.
//! - Rich multi-format strings.
//! - Outline groupings.
//! - Defined names.
//! - Autofilters.
//! - Worksheet Tables.
//! - Serde serialization support.
//! - Support for macros.
//! - Memory optimization mode for writing large files.
//!
//! [`XlsxWriter`]: https://xlsxwriter.readthedocs.io/index.html
//!
//! # Table of Contents
//!
//! - [`Tutorial`](crate::tutorial): A getting started and tutorial guide.
//! - [`Cookbook`](crate::cookbook): Examples of using `rust_xlsxwriter`.
//!
//! <p>
//!
//! - [`Workbook`]: The entry point for creating an Excel workbook with
//!   worksheets.
//! - [`Working with Workbooks`](crate::workbook): A higher-level introduction
//!   to creating and working with workbooks.
//! </p>
//!
//! <p>
//!
//! - [`Worksheet`]: The main spreadsheet canvas for writing data and objects to
//!   a worksheet.
//! - [`Working with Worksheets`](crate::worksheet): A higher-level introduction
//!   to creating and working with worksheets.
//! </p>
//!
//! <p>
//!
//! - [`Chart`] struct: The interface for creating worksheet charts.
//! - [`Working with Charts`](crate::chart): A higher-level introduction to
//!   creating and using charts.
//! </p>
//!
//! <p>
//!
//! - [`Format`]: The interface for adding formatting to worksheets and other
//!   objects.
//! - [`Table`]: The interface for worksheet tables. Tables in Excel are a way
//!   of grouping a range of cells into a single entity that has common
//!   formatting or that can be referenced in formulas.
//! - [`Image`]: The interface for images used in worksheets.
//! - [`Conditional Formats`](crate::conditional_format): Working with
//!   conditional formatting in worksheets.
//! - [`DataValidation`]: Working with data validation in worksheets.
//! - [`Note`]: Adding Notes to worksheet cells.
//! - [`Shape`]: Adding Textbox shapes to worksheets.
//! - [`Macros`](crate::macros): Working with Macros.
//! - [`Sparklines`](crate::sparkline): Working with Sparklines.
//! - [`ExcelDateTime`]: A type to represent dates and times in Excel format.
//! - [`Formula`]: A type for Excel formulas.
//! - [`Url`]: A type for URLs/Hyperlinks used in worksheets.
//! - [`DocProperties`]: The interface used to create an object to represent
//!   document metadata properties.
//! </p>
//!
//! - [`Changelog`](crate::changelog): Release notes and changelog.
//! - [`Performance`](crate::performance): Performance characteristics of
//!   `rust_xlsxwriter`.
//!
//! Other external documentation:
//!
//! - [User Guide]: Working with the `rust_xlsxwriter` library.
//! - [Roadmap of Planned Features].
//!
//! [User Guide]: https://rustxlsxwriter.github.io/index.html
//! [Roadmap of Planned Features]:
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
//! See the [`Cookbook`](crate::cookbook) for more examples.
//!
//!
//! # Motivation
//!
//! The `rust_xlsxwriter` crate was designed and implemented based around the
//! following design considerations:
//!
//! - **Fidelity with the Excel file format**. The library uses its own XML
//!   writer module in order to be as close as possible to the format created by
//!   Excel. It also contains a test suite of over 1,000 tests that compare
//!   generated files with those created by Excel. This has the advantage that
//!   it rarely creates a file that isn't compatible with Excel, and also that
//!   it is easy to debug and maintain because it can be compared with an Excel
//!   sample file using a simple diff.
//! - **Performance**. The library is designed to be as fast and efficient as
//!   possible. It also supports a constant memory mode for writing large files,
//!   which keeps memory usage to a minimum.
//! - **Comprehensive documentation**. In addition to the API documentation, the
//!   library has extensive user guides, a tutorial, and a cookbook of examples.
//!   It also includes images of Excel with the output of most of the example
//!   code.
//! - **Feature richness**. The library supports a wide range of Excel features,
//!   including charts, conditional formatting, data validation, rich text,
//!   hyperlinks, images, and even sparklines. It also supports new Excel 365
//!   features like dynamic arrays and spill ranges.
//! - **Write only**. The library only supports writing Excel files, and not
//!   reading or modifying them. This allows it to focus on doing one task as
//!   comprehensively as possible.
//! - **A family of libraries**. The `rust_xlsxwriter` library has sister
//!   libraries written in C ([libxlsxwriter]), Python ([XlsxWriter]), and Perl
//!   ([Excel::Writer::XLSX]), by the same author. Bug fixes and improvements in
//!   one get transferred to the others.
//! - **No FAQ section**. The Rust implementation seeks to avoid some of the
//!   required workarounds and API mistakes of the other language variants. For
//!   example, it has a `save()` function, automatic handling of dynamic
//!   functions, a much more transparent Autofilter implementation, and was the
//!   first version to have Autofit.
//!
//! [XlsxWriter]: https://xlsxwriter.readthedocs.io/index.html
//! [libxlsxwriter]: https://libxlsxwriter.github.io
//! [Excel::Writer::XLSX]:
//!     https://metacpan.org/dist/Excel-Writer-XLSX/view/lib/Excel/Writer/XLSX.pm
//!
//!
//! # Performance
//!
//! As mentioned above the `rust_xlsxwriter` library has sister libraries
//! written natively in C, Python, and Perl.
//!
//! A relative performance comparison between the C, Rust, and Python versions
//! is shown below. The Perl performance is similar to the Python library, so it
//! has been omitted.
//!
//! | Library                       | Relative to C | Relative to Rust |
//! |-------------------------------|---------------|------------------|
//! | C/libxlsxwriter               | 1.00          |                  |
//! | `rust_xlsxwriter`             | 1.14          | 1.00             |
//! | Python/XlsxWriter             | 4.36          | 3.81             |
//!
//! <br>
//!
//! The C version is the fastest: it is 1.14 times faster than the Rust version
//! and 4.36 times faster than the Python version. The Rust version is 3.81
//! times faster than the Python version.
//!
//! See the [Performance] section for more details.
//!
//! [Performance]: performance/index.html
//!
//!
//! # Crate Features
//!
//! The following is a list of the features supported by the `rust_xlsxwriter`
//! crate.
//!
//! **Default**
//!
//! - `default`: This includes all the standard functionality. The only
//!   dependency is the `zip` crate.
//!
//! **Optional features**
//!
//! These are all off by default.
//!
//! - `constant_memory`: Keeps memory usage to a minimum when writing large
//!   files. See [Constant Memory
//!   Mode](performance/index.html#constant-memory-mode).
//! - `serde`: Adds support for Serde serialization.
//! - `chrono`: Adds support for Chrono date/time types to the API. See
//!   [`IntoExcelDateTime`].
//! - `jiff`: Adds support for Jiff date/time types to the API. See
//!   [`IntoExcelDateTime`].
//! - `zlib`: Improves performance of the `zlib` crate but adds a dependency on
//!   zlib and a C compiler. This can be up to 1.5 times faster for large files.
//! - `polars`: Adds support for mapping between `PolarsError` and
//!   `rust_xlsxwriter::XlsxError` to make code that handles both types of
//!   errors easier to write. See also
//!   [`polars_excel_writer`](https://crates.io/crates/polars_excel_writer).
//! - `wasm`: Adds a dependency on `js-sys` and `wasm-bindgen` to allow
//!   compilation for wasm/JavaScript targets. See also
//!   [wasm-xlsxwriter](https://github.com/estie-inc/wasm-xlsxwriter).
//! - `rust_decimal`: Adds support for writing the
//!   [`rust_decimal`](https://crates.io/crates/rust_decimal) `Decimal` type
//!   with `Worksheet::write()`, provided it can be represented by [`f64`].
//! - `ryu`: Adds a dependency on `ryu`. This speeds up writing numeric
//!   worksheet cells for large data files. It gives a performance boost for
//!   more than 300,000 numeric cells and can be up to 30% faster than the
//!   default number formatting for 5,000,000 numeric cells.
//!
//! A `rust_xlsxwriter` feature can be enabled in your `Cargo.toml` file as
//! follows:
//!
//! ```bash
//! cargo add rust_xlsxwriter -F constant_memory
//! ```
//!
mod app;
mod button;
mod color;
mod comment;
mod content_types;
mod core;
mod custom;
mod data_validation;
mod datetime;
mod drawing;
mod error;
mod feature_property_bag;
mod filter;
mod format;
mod formula;
mod image;
mod metadata;
mod note;
mod packager;
mod properties;
mod protection;
mod relationship;
mod rich_value;
mod rich_value_rel;
mod rich_value_structure;
mod rich_value_types;
mod shape;
mod shared_strings;
mod shared_strings_table;
mod styles;
mod table;
mod theme;
mod url;
mod vml;
mod xmlwriter;

#[cfg(feature = "serde")]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub mod serializer;

pub mod changelog;
pub mod chart;
pub mod conditional_format;
pub mod cookbook;
pub mod macros;
pub mod performance;
pub mod sparkline;
pub mod tutorial;
pub mod utility;
pub mod workbook;
pub mod worksheet;

#[cfg(test)]
mod test_functions;

// Re-export the public APIs.
pub use button::*;
pub use color::*;
pub use data_validation::*;
pub use datetime::*;
pub use error::*;
pub use filter::*;
pub use format::*;
pub use formula::*;
pub use image::*;
pub use note::*;
pub use properties::*;
pub use protection::*;
pub use shape::*;
pub use table::*;
pub use url::*;

#[doc(hidden)]
pub use chart::*;

#[doc(hidden)]
pub use comment::*;

#[doc(hidden)]
pub use conditional_format::*;

#[doc(hidden)]
pub use sparkline::*;

#[doc(hidden)]
pub use worksheet::*;

#[doc(hidden)]
pub use workbook::*;

#[doc(hidden)]
pub use utility::*;

#[cfg(feature = "serde")]
#[doc(hidden)]
pub use serializer::*;

#[cfg(feature = "serde")]
extern crate rust_xlsxwriter_derive;

#[cfg(feature = "serde")]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub use rust_xlsxwriter_derive::XlsxSerialize;
