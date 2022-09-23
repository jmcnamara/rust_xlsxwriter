// Entry point for rust_xlsxwriter library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The `rust_xlsxwriter` library is a rust library for writing Excel files in
//! the XLSX format.
//!
//! <img
//! src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/demo.png">
//!
//! Rust_xlsxwriter is a rust library that can be used to write text, numbers,
//! dates and formulas to multiple worksheets in a new Excel 2007+ XLSX file. It
//! has a focus on performance and on fidelity with file format created by
//! Excel. It cannot be used to modify an existing file.
//!
//! # Examples
//!
//! Sample code to generate the Excel file shown above.
//!
//! ```rust
//! # // This code is available in examples/app_demo.rs
//! #
//! use chrono::NaiveDate;
//! use rust_xlsxwriter::{Format, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file.
//!     let mut workbook = Workbook::new("demo.xlsx");
//!
//!     // Create some formats to use in the worksheet.
//!     let bold_format = Format::new().set_bold();
//!     let decimal_format = Format::new().set_num_format("0.000");
//!     let date_format = Format::new().set_num_format("yyyy-mm-dd");
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Set the column width for clarity.
//!     worksheet.set_column_width(0, 15)?;
//!
//!     // Write a string without formatting.
//!     worksheet.write_string_only(0, 0, "Hello")?;
//!
//!     // Write a string with the bold format defined above.
//!     worksheet.write_string(1, 0, "World", &bold_format)?;
//!
//!     // Write some numbers.
//!     worksheet.write_number_only(2, 0, 1)?;
//!     worksheet.write_number_only(3, 0, 2.34)?;
//!
//!     // Write a number with formatting.
//!     worksheet.write_number(4, 0, 3.00, &decimal_format)?;
//!
//!     // Write a formula.
//!     worksheet.write_formula_only(5, 0, "=SIN(PI()/4)")?;
//!
//!     // Write the date .
//!     let date = NaiveDate::from_ymd(2023, 1, 25);
//!     worksheet.write_date(6, 0, date, &date_format)?;
//!
//!     workbook.close()?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Rust_xlsxwriter is a port of the [XlsxWriter] Python module by the same
//! author. Feature porting is a work in progress. The currently supported
//! features are:
//!
//! - Support for writing all basic Excel data types.
//! - Full cell formatting support.
//!
//! [XlsxWriter]: https://xlsxwriter.readthedocs.io/index.html
//!
//!
//! # Getting started.
//!
//! Create a new rust command-line application as follows:
//!
//! ```bash
//! cargo new hello-xlsx
//! ```
//!
//! This will create a directory like the following:
//!
//! ```bash
//! hello-xlsx/
//! ├── Cargo.toml
//! └── src
//!     └── main.rs
//! ```
//!
//! Edit the Cargo.toml file and add the following `rust_xlsxwriter` dependency
//! so the file looks like below. Note, `rust_xlsxwriter` adds new features
//! regularly do make sure you use the latest version.
//!
//!
//! ```yaml
//! [package]
//! name = "hello-xlsx"
//! version = "0.1.0"
//! edition = "2021"
//!
//! [dependencies]
//! rust_xlsxwriter = "0.2.0"
//! ```
//!
//! Modify the main.rs file so it looks like this:
//!
//! ```rust
//! # // This code is available in examples/app_hello_world.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file.
//!     let mut workbook = Workbook::new("hello.xlsx");
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Write a string to cell (0, 0) = A1.
//!     worksheet.write_string_only(0, 0, "Hello")?;
//!
//!     // Write a number to cell (1, 0) = A2.
//!     worksheet.write_number_only(1, 0, 12345)?;
//!
//!     // Close the file.
//!     workbook.close()?;
//!
//!     Ok(())
//! }
//! ```
//!
//! The run the application as follows:
//!
//! ```bash
//! cargo run
//! ```
//!
//! This will create an output file called `hello.xlsx` which should look
//! something like this:
//!
//! <img
//! src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/hello.png">
//!
//! # Tutorial
//!
//! For something more ambitious than a hello world application we will use
//! `rust_xlsxwriter` to create a spreadsheet and add some monthly expenses.
//!
//! ```ignore
//!     let expenses = vec![
//!         ("Rent", 2000),
//!         ("Gas", 200),
//!         ("Food", 500),
//!         ("Gym", 100),
//!     ];
//! ```
//!
//! To do that we might start with a simple program like the following:
//!
//! ```
//! # // This code is available in examples/app_tutorial1.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Some sample data we want to write to a spreadsheet.
//!     let expenses = vec![("Rent", 2000), ("Gas", 200), ("Food", 500), ("Gym", 100)];
//!
//!     // Create a new Excel file.
//!     let mut workbook = Workbook::new("tutorial1.xlsx");
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Iterate over the data and write it out row by row.
//!     let mut row = 0;
//!     for expense in expenses.iter() {
//!         worksheet.write_string_only(row, 0, expense.0)?;
//!         worksheet.write_number_only(row, 1, expense.1)?;
//!         row += 1;
//!     }
//!
//!     // Write a total using a formula.
//!     worksheet.write_string_only(row, 0, "Total")?;
//!     worksheet.write_formula_only(row, 1, "=SUM(B1:B4)")?;
//!
//!     // Close the file.
//!     workbook.close()?;
//!
//!     Ok(())
//! }
//! ```
//!
//! If we run this program we should get a spreadsheet that looks like this:
//!
//! <img
//! src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/tutorial1.png">
//!
//! The first step is to create a new workbook object using the
//! [`Workbook`](Workbook) constructor. [`Workbook::new`](Workbook::new) takes
//! one argument which is the filename that we want to create:
//!
//! ```ignore
//!     let mut workbook = Workbook::new("tutorial1.xlsx");
//! ```
//!
//! **Note**, `rust_xlsxwriter` can only create new files. It cannot read or
//! modify existing files.
//!
//! The workbook object is then used to add a new worksheet via the
//! [`add_worksheet`](Workbook::add_worksheet) method:
//!
//! ```ignore
//!     let worksheet = workbook.add_worksheet();
//! ```
//! The worksheet will have a standard Excel name, in this case "Sheet1". You
//! can specify the worksheet name using the
//! [`worksheet.set_name()`](Worksheet::set_name) method.
//!
//! We then iterate over the data and use some the
//! [`worksheet.write_string_only()`](Worksheet::write_string_only) and
//! [`worksheet.write_number_only()`](Worksheet::write_number_only) methods to
//! write each row of our data:
//!
//! ```ignore
//!         worksheet.write_string_only(row, 0, expense.0)?;
//!         worksheet.write_number_only(row, 1, expense.1)?;
//! ```
//!
//! The `_only` part of the method names refers to the fact that the data type
//! is written without any formatting. We will see how to add formatting
//! shortly.
//!
//! Throughout `rust_xlsxwriter`, rows and columns are zero indexed. So, for
//! example, the first cell in a worksheet, `A1`, is `(0, 0)`.
//!
//! We then add a formula to calculate the total of the items in the second
//! column:
//!
//! ```ignore
//!     worksheet.write_formula_only(row, 1, "=SUM(B1:B4)")?;
//! ```
//!
//! Finally, we close the Excel file via the
//! [`workbook.close()`](Workbook::close) method:
//!
//! ```ignore
//!     workbook.close()?;
//! ```
//!
//! The previous program converted the required data into an Excel file but it
//! looked a little bare. In order to make the information clearer we can add
//! some simple formatting, like this:
//!
//! <img
//! src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/tutorial2.png">
//!
//! The differences here are that we have added "Item" and "Cost" column headers
//! in a bold font, we have formatted the currency in the second column and we
//! have made the "Total" string bold.
//!
//! To do this programmatically we can extend our code as follows:
//!
//! ```
//! # // This code is available in examples/app_tutorial2.rs
//! #
//! use rust_xlsxwriter::{Format, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Some sample data we want to write to a spreadsheet.
//!     let expenses = vec![("Rent", 2000), ("Gas", 200), ("Food", 500), ("Gym", 100)];
//!
//!     // Create a new Excel file.
//!     let mut workbook = Workbook::new("tutorial2.xlsx");
//!
//!     // Add a bold format to use to highlight cells.
//!     let bold = Format::new().set_bold();
//!
//!     // Add a number format for cells with money values.
//!     let money_format = Format::new().set_num_format("$#,##0");
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Write some column headers.
//!     worksheet.write_string(0, 0, "Item", &bold)?;
//!     worksheet.write_string(0, 1, "Cost", &bold)?;
//!
//!     // Iterate over the data and write it out row by row.
//!     let mut row = 1;
//!     for expense in expenses.iter() {
//!         worksheet.write_string_only(row, 0, expense.0)?;
//!         worksheet.write_number(row, 1, expense.1, &money_format)?;
//!         row += 1;
//!     }
//!
//!     // Write a total using a formula.
//!     worksheet.write_string(row, 0, "Total", &bold)?;
//!     worksheet.write_formula(row, 1, "=SUM(B2:B5)", &money_format)?;
//!
//!     // Close the file.
//!     workbook.close()?;
//!
//!     Ok(())
//! }
//! ```
//!
//! The main difference between this and the previous program is that we have
//! added two [Format](Format) objects that we can use to format cells in the
//! spreadsheet.
//!
//! Format objects represent all of the formatting properties that can be
//! applied to a cell in Excel such as fonts, number formatting, colors and
//! borders. This is explained in more detail in the [Format](Format) struct
//! documentation.
//!
//! For now we will avoid getting into the details of Format and just use a
//! limited amount of the its functionality to add some simple formatting:
//!
//! ```ignore
//!     // Add a bold format to use to highlight cells.
//!     let bold = Format::new().set_bold();
//!
//!     // Add a number format for cells with money values.
//!     let money_format = Format::new().set_num_format("$#,##0");
//! ```
//!
//! We can use these formats in worksheet methods that support formatting such
//! as the [`worksheet.write_string()`](Worksheet::write_string) and
//! [`worksheet.write_number()`](Worksheet::write_number) methods which can
//! write data and formatting together.
//!
//! Let's extend the program a little bit more to add some dates to the data:
//!
//! ```ignore
//!     let expenses = vec![
//!         ("Rent", 2000, "2022-09-01"),
//!         ("Gas", 200, "2022-09-05"),
//!         ("Food", 500, "2022-09-21"),
//!         ("Gym", 100, "2022-09-28"),
//!     ];
//! ```
//!
//! The corresponding spreadsheet will look like this:
//!
//! <img
//! src="https://github.com/jmcnamara/rust_xlsxwriter/raw/main/examples/images/tutorial3.png">
//!
//! The differences here are that we have added a "Date" column with formatting
//! and made that column a little wider to accommodate the dates.
//!
//! To do this we can extend our program as follows:
//!
//! ```
//! # // This code is available in examples/app_tutorial3.rs
//! #
//! use chrono::NaiveDate;
//! use rust_xlsxwriter::{Format, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Some sample data we want to write to a spreadsheet.
//!     let expenses = vec![
//!         ("Rent", 2000, "2022-09-01"),
//!         ("Gas", 200, "2022-09-05"),
//!         ("Food", 500, "2022-09-21"),
//!         ("Gym", 100, "2022-09-28"),
//!     ];
//!
//!     // Create a new Excel file.
//!     let mut workbook = Workbook::new("tutorial3.xlsx");
//!
//!     // Add a bold format to use to highlight cells.
//!     let bold = Format::new().set_bold();
//!
//!     // Add a number format for cells with money values.
//!     let money_format = Format::new().set_num_format("$#,##0");
//!
//!     // Add a number format for cells with dates.
//!     let date_format = Format::new().set_num_format("d mmm yyyy");
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Write some column headers.
//!     worksheet.write_string(0, 0, "Item", &bold)?;
//!     worksheet.write_string(0, 1, "Cost", &bold)?;
//!     worksheet.write_string(0, 2, "Date", &bold)?;
//!
//!     // Adjust the date column width for clarity.
//!     worksheet.set_column_width(2, 15)?;
//!
//!     // Iterate over the data and write it out row by row.
//!     let mut row = 1;
//!     for expense in expenses.iter() {
//!         worksheet.write_string_only(row, 0, expense.0)?;
//!         worksheet.write_number(row, 1, expense.1, &money_format)?;
//!
//!         let date = NaiveDate::parse_from_str(expense.2, "%Y-%m-%d").unwrap();
//!         worksheet.write_date(row, 2, date, &date_format)?;
//!
//!         row += 1;
//!     }
//!
//!     // Write a total using a formula.
//!     worksheet.write_string(row, 0, "Total", &bold)?;
//!     worksheet.write_formula(row, 1, "=SUM(B2:B5)", &money_format)?;
//!
//!     // Close the file.
//!     workbook.close()?;
//!
//!     Ok(())
//! }
//!```
//!
//! The main difference between this and the previous program is that we have
//! added handling for dates.
//!
//! Dates and times in Excel are floating point numbers that have a number
//! format applied to display them in the correct format. In order to handle
//! dates and times with `rust_xlsxwriter` we create them as
//! [`chrono::NaiveDateTime`], [`chrono::NaiveDate`] or [`chrono::NaiveTime`]
//! instances and format them with an Excel number format.
//!
//! In the example above we create the NaiveDates from the date strings in our
//! input data and then format it, in Excel, with a number format.
//!
//! ```ignore
//!     // Add a number format for cells with dates.
//!     let date_format = Format::new().set_num_format("d mmm yyyy");
//! ```
//!
//! The final addition to our program is the make the "Date" column wider for
//! clarity using the
//! [`worksheet.set_column_width()`](Worksheet::set_column_width) method.
//!
//! That completes the tutorial section. For more information about the
//! available Structs and associated methods see the documentation for
//! [`Workbook`], [`Worksheet`] and [`Format`].
//!
mod app;
mod content_types;
mod core;
mod error;
mod format;
mod packager;
mod relationship;
mod shared_strings;
mod shared_strings_table;
mod styles;
mod test_functions;
mod theme;
mod utility;
mod workbook;
mod worksheet;
mod xmlwriter;

// Re-export the public APIs.
pub use error::*;
pub use format::*;
pub use workbook::*;
pub use worksheet::*;
