// worksheet - A module for creating the Excel sheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! # Working with Worksheets
//!
//! The [`Worksheet`] struct represents an Excel worksheet. It handles
//! operations such as writing data to cells or formatting the worksheet layout.
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
//! For more details on the Worksheet APIs, see the [`Worksheet`]
//! documentation and the sections below.
//!
//! # Contents
//!
//! - [Creating worksheets](#creating-worksheets)
//!   - [Working with `add_worksheet()` and the borrow
//!     checker](#working-with-add_worksheet-and-the-borrow-checker)
//!   - [Working with `Worksheet::new()` and the borrow
//!     checker](#working-with-worksheetnew-and-the-borrow-checker)
//!   - [Using `add_worksheet()` versus
//!     `Worksheet::new()`](#using-add_worksheet-versus--worksheetnew)
//! - [Chartsheets](#chartsheets)
//! - [Page Setup](#page-setup)
//!   - [Page Setup - Page](#page-setup---page)
//!   - [Page Setup - Margins](#page-setup---margins)
//!   - [Page Setup - Header/Footer](#page-setup---headerfooter)
//!   - [Page Setup - Sheet](#page-setup---sheet)
//! - [Cell formatting](#cell-formatting)
//! - [Adding Headers and Footers](#adding-headers-and-footers)
//! - [Autofitting column widths](#autofitting-column-widths)
//! - [Working with worksheet tabs](#working-with-worksheet-tabs)
//!   - [Worksheet names](#worksheet-names)
//!   - [Setting the active worksheet](#setting-the-active-worksheet)
//!   - [Setting worksheet tab colors](#setting-worksheet-tab-colors)
//!   - [Hiding worksheets](#hiding-worksheets)
//!   - [Selecting worksheets](#selecting-worksheets)
//! - [Grouping and outlining data](#grouping-and-outlining-data)
//! - [Worksheet protection](#worksheet-protection)
//!   - [Setting a protection password](#setting-a-protection-password)
//!   - [Choosing which worksheet elements to
//!     protect](#choosing-which-worksheet-elements-to-protect)
//!   - [Workbook protection](#workbook-protection)
//!   - [Read-only workbook](#read-only-workbook)
//!
//!
//! # Creating worksheets
//!
//! There are two ways to create a Worksheet object with `rust_xlsxwriter`:
//! using the [`Workbook::add_worksheet()`] method or the
//! [`Worksheet::new()`] constructor.
//!
//! The first method ties the worksheet to a workbook object that will
//! automatically write the worksheet when the file is saved, whereas the second
//! method creates a worksheet that is independent of a workbook. The second
//! method has the advantage of keeping the worksheet free of the workbook
//! borrow checking until needed, as explained below.
//!
//!
//! ## Working with `add_worksheet()` and the borrow checker
//!
//! Due to borrow checking rules you can only have one active reference to a
//! worksheet object created by [`Workbook::add_worksheet()`] since that method
//! returns a mutable reference to an element of an internal vector.
//!
//! For a workbook with multiple worksheets this restriction is generally
//! workable if you can create and use the worksheets sequentially since you
//! will only need to have one reference at any one time.
//!
//! However, if you can’t structure your code to work sequentially, you can
//! get a reference to a previously created worksheet using
//! [`Workbook::worksheet_from_index()`]. The standard borrow checking rules
//! still apply so you will have to give up ownership of any other worksheet
//! reference prior to calling this method.
//!
//! For example:
//!
//! ```
//! # // This code is available in examples/doc_workbook_worksheet_from_index.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     // Start with a reference to worksheet1.
//!     let mut worksheet1 = workbook.add_worksheet();
//!     worksheet1.write_string(0, 0, "Hello")?;
//!
//!     // If we don't try to use the workbook1 reference again we can
//!     // just switch to using a reference to worksheet2.
//!     let mut worksheet2 = workbook.add_worksheet();
//!     worksheet2.write_string(0, 0, "Hello")?;
//!
//!     // Stop using worksheet2 and move back to worksheet1.
//!     worksheet1 = workbook.worksheet_from_index(0)?;
//!     worksheet1.write_string(1, 0, "Sheet1")?;
//!
//!     // Stop using worksheet1 and move back to worksheet2.
//!     worksheet2 = workbook.worksheet_from_index(1)?;
//!     worksheet2.write_string(1, 0, "Sheet2")?;
//! #
//! #     workbook.save("workbook.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! You can also use [`Workbook::worksheet_from_name()`] to do something similar
//! using the worksheet names:
//!
//! ```
//! # // This code is available in examples/doc_workbook_worksheet_from_name.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     // Start with a reference to worksheet1.
//!     let mut worksheet1 = workbook.add_worksheet();
//!     let name1 = worksheet1.name(); // "Sheet1"
//!     worksheet1.write_string(0, 0, "Hello")?;
//!
//!     // If we don't try to use the workbook1 reference again we can
//!     // just switch to using a reference to worksheet2.
//!     let mut worksheet2 = workbook.add_worksheet().set_name("Data")?;
//!     let name2 = worksheet2.name();
//!     worksheet2.write_string(0, 0, "Hello")?;
//!
//!     // Stop using worksheet2 and move back to worksheet1.
//!     worksheet1 = workbook.worksheet_from_name(&name1)?;
//!     worksheet1.write_string(1, 0, "Sheet1")?;
//!
//!     // Stop using worksheet1 and move back to worksheet2.
//!     worksheet2 = workbook.worksheet_from_name(&name2)?;
//!     worksheet2.write_string(1, 0, "Sheet2")?;
//! #
//! #     workbook.save("workbook.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Additionally, you can use [`Workbook::worksheets_mut()`] to get a reference
//! to the vector that holds the worksheets. This can be used, for example, to
//! iterate over all the worksheets in the workbook:
//!
//! ```
//! # // This code is available in examples/doc_workbook_worksheets_mut.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     // Add three worksheets to the workbook.
//!     let _worksheet1 = workbook.add_worksheet();
//!     let _worksheet2 = workbook.add_worksheet();
//!     let _worksheet3 = workbook.add_worksheet();
//!
//!     // Write the same data to all three worksheets.
//!     for worksheet in workbook.worksheets_mut() {
//!         worksheet.write_string(0, 0, "Hello")?;
//!         worksheet.write_number(1, 0, 12345)?;
//!     }
//!
//!     // If you are careful you can use standard slice operations.
//!     workbook.worksheets_mut().swap(0, 1);
//! #
//! #     workbook.save("workbook.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Note the use of the [slice] operation which gives the following:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/workbook_worksheets_mut.png">
//!
//! All three worksheets in the output file have the same data and `Sheet2` and
//! `Sheet1` have swapped position, as can be seen from the image.
//!
//! ## Working with `Worksheet::new()` and the borrow checker
//!
//! As outlined above it is also possible to create a Worksheet object via
//! [`Worksheet::new()`], as you would expect. Since this type of Worksheet
//! instance isn't tied to a Workbook it isn't subject to the additional borrow
//! checking rules that entails.
//!
//! As such you can create and work with several Worksheet instances and then
//! add them to the Workbook when you are finished via the
//! [`Workbook::push_worksheet()`] method:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_new.rs
//! #
//! # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//!     // Create a new workbook.
//!     let mut workbook = Workbook::new();
//!
//!     // Create new worksheets.
//!     let mut worksheet1 = Worksheet::new();
//!     let mut worksheet2 = Worksheet::new();
//!
//!     // Use the first workbook.
//!     worksheet1.write_string(0, 0, "Hello")?;
//!     worksheet1.write_string(1, 0, "Sheet1")?;
//!
//!     // Use the second workbook.
//!     worksheet2.write_string(0, 0, "Hello")?;
//!     worksheet2.write_string(1, 0, "Sheet2")?;
//!
//!     // Add the worksheets to the workbook.
//!     workbook.push_worksheet(worksheet1);
//!     workbook.push_worksheet(worksheet2);
//!
//!     // Save the workbook.
//!     workbook.save("worksheets.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! ## Using `add_worksheet()` versus  `Worksheet::new()`
//!
//! Since there are two ways of doing, effectively, the same thing you will
//! probably wonder which is best.
//!
//! The author mainly uses `add_worksheet()` and most of the library and
//! documentation examples are structured to work with that method. In addition,
//! the Excel xlsx file format has very tight coupling between its
//! sub-components and it is possible that some future `rust_xlsxwriter`
//! functionality will require Worksheets and other Workbook objects to be
//! `registered` with a parent Workbook in order for them to work. However,
//! there aren't currently any features like that, and the author will seek to
//! avoid them as much as possible.
//!
//! One common use case that works better with `Worksheet::new()` and
//! `Workbook::push_worksheet()` is creating worksheets to run in a
//! parallelized/async mode. However, it is worth noting that there isn't a
//! guaranteed performance benefit from creating and working with worksheets in
//! parallelized/async mode since the main overhead comes from **writing** the
//! worksheets which will occur after the worksheets are joined back to the main
//! workbook `save()` thread. In addition `rust_xlsxwriter` already parallelizes
//! the writing of worksheets as much as possible.
//!
//! [`Workbook::add_worksheet()`]: crate::Workbook::add_worksheet
//! [`Workbook::worksheets_mut()`]: crate::Workbook::worksheets_mut
//! [`Workbook::push_worksheet()`]: crate::Workbook::push_worksheet
//! [`Workbook::worksheet_from_name()`]: crate::Workbook::worksheet_from_name
//! [`Workbook::worksheet_from_index()`]: crate::Workbook::worksheet_from_index
//!
//!
//! # Chartsheets
//!
//! A Chartsheet in Excel is a specialized type of worksheet that doesn't have
//! cells but instead is used to display a single chart. It supports worksheet
//! display options such as headers and footers, margins, tab selection and
//! print properties but its primary function is to display a chart.
//!
//! <img src="https://rustxlsxwriter.github.io/images/chartsheet.png">
//!
//! A chartsheet can be created using the
//! [`Workbook::add_chartsheet()`](crate::Workbook::add_chartsheet) of
//! [`Worksheet::new_chartsheet()`] methods.
//!
//! The following is a simple chartsheet example that creates a workbook with a
//! worksheet and a chartsheet as shown in the image above.
//!
//! ```
//! # // This code is available in examples/doc_chartsheet.rs
//! #
//! use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Add some data for the chart.
//!     worksheet.write(0, 0, 10)?;
//!     worksheet.write(1, 0, 60)?;
//!     worksheet.write(2, 0, 30)?;
//!     worksheet.write(3, 0, 10)?;
//!     worksheet.write(4, 0, 50)?;
//!
//!     // Create a new chart.
//!     let mut chart = Chart::new(ChartType::Column);
//!
//!     // Add a data series using Excel formula syntax to describe the range.
//!     chart.add_series().set_values("Sheet1!$A$1:$A$5");
//!
//!     // Create a new chartsheet.
//!     let chartsheet = workbook.add_chartsheet();
//!
//!     // Add the chart to the chartsheet.
//!     chartsheet.insert_chart(0, 0, &chart)?;
//!
//!     // Save the file.
//!     workbook.save("chart.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! # Page Setup
//!
//! The sections below look at the elements of the Excel Page Setup dialog and
//! the equivalent `rust_xlsxwriter` methods.
//!
//! For more, general, information about the page setup options in Excel see the
//! Microsoft documentation on [Page Setup].
//!
//! [Page Setup]:
//!     https://support.microsoft.com/en-us/office/page-setup-71c20d94-b13e-48fd-9800-cedd1fec6da3
//!
//! ## Page Setup - Page
//!
//! The page Setup "Page" dialog looks like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/page_setup01.png">
//!
//! The equivalent `rust_xlsxwriter` methods are:
//!
//! 1. [`Worksheet::set_portrait()`]
//! 2. [`Worksheet::set_landscape()`]
//! 3. [`Worksheet::set_print_scale()`]
//! 4. [`Worksheet::set_print_fit_to_pages()`]
//! 5. [`Worksheet::set_print_first_page_number()`]
//!
//! Note, for [`Worksheet::set_print_fit_to_pages()`] a common requirement is to
//! fit the printed output to `n` pages wide but have the height be as long as
//! necessary. To achieve this set the `height` to 0:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_print_fit_to_pages.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Set the printed output to fit 1 page wide and as long as necessary.
//!     worksheet.set_print_fit_to_pages(1, 0);
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//!
//! ## Page Setup - Margins
//!
//! The page Setup "Margins" dialog looks like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/page_setup02.png">
//!
//! The equivalent `rust_xlsxwriter` methods are:
//!
//! 1. [`Worksheet::set_margins()`]
//! 2. [`Worksheet::set_print_center_horizontally()`]
//! 3. [`Worksheet::set_print_center_vertically()`]
//!
//!
//! ## Page Setup - Header/Footer
//!
//! The page Setup "Header/Footer" dialog looks like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/page_setup03.png">
//!
//! The equivalent `rust_xlsxwriter` methods are:
//!
//! 1. [`Worksheet::set_header()`]
//! 2. [`Worksheet::set_footer()`]
//! 3. [`Worksheet::set_header_footer_scale_with_doc()`]
//! 4. [`Worksheet::set_header_footer_align_with_page()`]
//!
//! Headers and footers are explained in more detail in a subsequent section
//! below on [Adding Headers and Footers](#adding-headers-and-footers).
//!
//! Note, the options for different first, odd and even pages are not supported
//! in `rust_xlsxwriter`.
//!
//! ## Page Setup - Sheet
//!
//! The page Setup "Sheet" dialog looks like this:
//!
//! <img src="https://rustxlsxwriter.github.io/images/page_setup04.png">
//!
//! The equivalent `rust_xlsxwriter` methods are:
//!
//! 1. [`Worksheet::set_print_area()`]
//! 2. [`Worksheet::set_repeat_rows()`]
//! 3. [`Worksheet::set_repeat_columns()`]
//! 4. [`Worksheet::set_print_gridlines()`]
//! 5. [`Worksheet::set_print_black_and_white()`]
//! 6. [`Worksheet::set_print_draft()`]
//! 7. [`Worksheet::set_print_headings()`]
//! 8. [`Worksheet::set_page_order()`]
//!
//!
//! # Cell formatting
//!
//! In Excel the data in a worksheet cell is comprised of a type, a value and a
//! format. When using `rust_xlsxwriter` the type is inferred and the value and
//! format are generally written at the same time using methods like
//! [`Worksheet::write_with_format()`]:
//!
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_range_format2.rs
//! #
//! use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Add a format.
//!     let border = Format::new().set_border(FormatBorder::Thin);
//!
//!     // Some data to write.
//!     let data = [
//!         [10, 11, 12, 13, 14],
//!         [20, 21, 22, 23, 24],
//!         [30, 31, 32, 33, 34],
//!     ];
//!
//!     // Write the data with formatting.
//!     for (row_num, col) in data.iter().enumerate() {
//!         for (col_num, cell) in col.iter().enumerate() {
//!             let row_num = row_num as u32 + 1;
//!             let col_num = col_num as u16 + 1;
//!             worksheet.write_with_format(row_num, col_num, *cell, &border)?;
//!         }
//!     }
//!
//!     workbook.save("worksheet.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Which produces an output file like this:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_set_range_format.png">
//!
//! However, it is sometimes easier to structure `rust_xlsxwriter` programs to
//! write the data first and then add the formatting. To do that you can make
//! use of the following worksheet methods:
//!
//! - [`Worksheet::set_cell_format()`]
//! - [`Worksheet::set_range_format()`]
//! - [`Worksheet::set_range_format_with_border()`]
//!
//! Here is an example with the same output as  the previous example where the
//! data and formatting are handled separately:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_range_format.rs
//! #
//! use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Add a format.
//!     let border = Format::new().set_border(FormatBorder::Thin);
//!
//!     // Write an array of data.
//!     let data = [
//!         [10, 11, 12, 13, 14],
//!         [20, 21, 22, 23, 24],
//!         [30, 31, 32, 33, 34],
//!     ];
//!     worksheet.write_row_matrix(1, 1, data)?;
//!
//!     // Add formatting to the cells.
//!     worksheet.set_range_format(1, 1, 3, 5, &border)?;
//!
//!     workbook.save("worksheet.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//! The methodology of separating the data and the formatting is particularly
//! useful when you need to add a border to a range, since that can require up
//! to 9 separate formats and tracking of the cell positions. Here is an example
//! with [`Worksheet::set_range_format_with_border()`]:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_range_format_with_border.rs
//! #
//! use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Add some formats.
//!     let inner_border = Format::new().set_border(FormatBorder::Thin);
//!     let outer_border = Format::new().set_border(FormatBorder::Double);
//!
//!     // Write an array of data.
//!     let data = [
//!         [10, 11, 12, 13, 14],
//!         [20, 21, 22, 23, 24],
//!         [30, 31, 32, 33, 34],
//!     ];
//!     worksheet.write_row_matrix(1, 1, data)?;
//!
//!     // Add formatting to the cells.
//!     worksheet.set_range_format_with_border(1, 1, 3, 5, &inner_border, &outer_border)?;
//!
//!     workbook.save("worksheet.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Which produces an output file like this:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_set_range_format_with_border.png">
//!
//!
//! For use cases where the cell formatting changes based on cell values
//! Conditional Formatting may be a better option (see [Working with Conditional
//! Formats](../conditional_format/index.html)). Additionally, formatting a
//! ranges of cells as a Worksheet Table may be a better option than simple cell
//! formatting (see the [`Table`] section of the documentation).
//!
//!
//! # Adding Headers and Footers
//!
//!
//! Headers and footers can be added to worksheets using the
//! [`Worksheet::set_header()`] and [`Worksheet::set_footer()`] methods.
//!
//! Headers and footers are generated using a string which is a combination of
//! plain text and optional control characters defined by Excel.
//!
//! The available control characters are:
//!
//! | Control              | Category      | Description           |
//! | -------------------- | ------------- | --------------------- |
//! | `&L`                 | Alignment     | Left                  |
//! | `&C`                 |               | Center                |
//! | `&R`                 |               | Right                 |
//! | `&[Page]`  or `&P`   | Information   | Page number           |
//! | `&[Pages]` or `&N`   |               | Total number of pages |
//! | `&[Date]`  or `&D`   |               | Date                  |
//! | `&[Time]`  or `&T`   |               | Time                  |
//! | `&[File]`  or `&F`   |               | File name             |
//! | `&[Tab]`   or `&A`   |               | Worksheet name        |
//! | `&[Path]`  or `&Z`   |               | Workbook path         |
//! | `&fontsize`          | Font          | Font size             |
//! | `&"font,style"`      |               | Font name and style   |
//! | `&U`                 |               | Single underline      |
//! | `&E`                 |               | Double underline      |
//! | `&S`                 |               | Strikethrough         |
//! | `&X`                 |               | Superscript           |
//! | `&Y`                 |               | Subscript             |
//! | `&[Picture]` or `&G` | Images        | Picture/image         |
//! | `&&`                 | Miscellaneous | Literal ampersand &   |
//!
//! Some of the placeholder variables have a long version like `&[Page]` and a
//! short version like `&P`. The longer version is displayed in the Excel
//! interface but the shorter version is the way that it is stored in the file
//! format. Either version is okay since `rust_xlsxwriter` will translate as
//! required.
//!
//! Headers and footers have 3 edit areas to the left, center and right. Text
//! can be aligned to these areas by prefixing the text with the control
//! characters `&L`, `&C` and `&R`.
//!
//! For example:
//!
//! ```text
//! worksheet.set_header("&LHello");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header01.png">
//!
//!
//! ```text
//! worksheet.set_header("&CHello");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header02.png">
//!
//! ```text
//! worksheet.set_header("&RHello");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header03.png">
//!
//! You can also have text in each of the alignment areas:
//!
//! ```text
//! worksheet.set_header("&LCiao&CBello&RCielo");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header04.png">
//!
//! The information control characters act as variables/templates that Excel
//! will update/expand as the workbook or worksheet changes.
//!
//! ```text
//! worksheet.set_header("&CPage &[Page] of &[Pages]");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header05.png">
//!
//!
//! Times and dates are in the user's default format:
//!
//! ```text
//! worksheet.set_header("&CUpdated at &[Time]");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header06.png">
//!
//! To insert an image in use `&[Picture]` or `&G`. You will also need to use
//! [`Worksheet::set_header_image()`] to set the corresponding image:
//!
//! ```text
//! let image = Image::new("examples/watermark.png")?;
//!
//! worksheet.set_header("&C&[Picture]");
//! worksheet.set_header_image(&image, XlsxImagePosition::Center);
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header10.png">
//!
//! To include a single literal ampersand `&` in a header or footer you should
//! use a double ampersand `&&`:
//!
//! ```text
//! worksheet.set_header("&CCuriouser && Curiouser - Attorneys at Law");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header07.png">
//!
//! You can specify the font size of a section of the text by prefixing it with
//! the control character `&n` where `n` is the font size:
//!
//! ```text
//! worksheet.set_header("&C&20Big Hello");
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header08.png">
//!
//!
//! You can specify the font of a section of the text by prefixing it with the
//! control sequence `&"font,style"` where `fontname` is a font name such as
//! Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
//! "Courier New" or "Times New Roman" and `style` is one of the standard
//! Windows font descriptions like “Regular”, “Italic”, “Bold” or “Bold Italic”:
//!
//! ```text
//! worksheet.set_header(r#"&C&"Courier New,Bold Italic"Hello"#);
//! ```
//!
//! <img src="https://rustxlsxwriter.github.io/images/header09.png">
//!
//! It is possible to combine all of these features together to create complex
//! headers and footers. If you set up a complex header in Excel you can
//! transfer it to `rust_xlsxwriter` by inspecting the string in the Excel file.
//! For example the following shows how unzip and grep the Excel XML sub-files
//! on a Linux system. The example uses libxml's xmllint to format the XML for
//! clarity:
//!
//! ```bash
//! $ unzip myfile.xlsm -d myfile
//! $ xmllint --format `find myfile -name "*.xml" | xargs` | \
//!     egrep "Header|Footer" | sed 's/&amp;/\&/g'
//!
//!  <headerFooter scaleWithDoc="0">
//!    <oddHeader>&L&P</oddHeader>
//!  </headerFooter>
//! ```
//!
//! **Note**: Excel requires that the header or footer string be less than 256
//! characters, including the control characters. Strings longer than this will
//! not be written, and a warning will be output.
//!
//! # Autofitting column widths
//!
//! Column widths in a `rust_xlsxwriter` worksheet can be adjusted automatically
//! using the [`Worksheet::autofit()`] method.
//!
//! There is no option in the xlsx file format that can be used to say "autofit
//! columns on loading". Auto-fitting of columns is something that Excel does at
//! runtime when it has access to all of the worksheet information as well as
//! the Windows functions for calculating display areas based on fonts and
//! formatting.
//!
//! As such [`Worksheet::autofit()`] simulates this behavior by calculating
//! string widths using metrics taken from Excel. This isn't perfect but for
//! most cases it should be sufficient and if not you can set your own widths,
//! see below.
//!
//! The `Worksheet::autofit()` method ignores columns that already have an
//! explicit column width set via [`Worksheet::set_column_width()`] or
//! [`Worksheet::set_column_width_pixels()`] if it is greater than the
//! calculated maximum width. Alternatively, calling these methods after
//! `Worksheet::autofit()` will override the autofit value.
//!
//! **Note**, `Worksheet::autofit()` iterates through all the cells in a
//! worksheet that have been populated with data and performs a length
//! calculation on each one, so it can have a performance overhead for larger
//! worksheets.
//!
//!
//! # Working with worksheet tabs
//!
//! Worksheet tabs in Excel allow the user to differentiate between different
//! worksheets.
//!
//! Worksheets in a workbook can be highlighted via the tab name, color,
//! position or the fact that it is active when the user opens the workbook.
//!
//! <img src="https://rustxlsxwriter.github.io/images/worksheet_tabs.png">
//!
//! The `rust_xlsxwriter` library provides a number of methods, explained below,
//! to emulate this functionality.
//!
//! ## Worksheet names
//!
//! The worksheet name can be set with [`Worksheet::set_name()`]:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_name.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     let _worksheet1 = workbook.add_worksheet(); // Defaults to Sheet1
//!     let _worksheet2 = workbook.add_worksheet().set_name("Foglio2");
//!     let _worksheet3 = workbook.add_worksheet().set_name("Data");
//!     let _worksheet4 = workbook.add_worksheet(); // Defaults to Sheet4
//! #
//! #     workbook.save("worksheets.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which gives the following output:
//!
//! <img src="https://rustxlsxwriter.github.io/images/worksheet_set_name.png">
//!
//! Excel applies various rules to worksheet names such as:
//!
//! * The name must be less than 32 characters.
//! * The name cannot be blank.
//! * The name cannot contain any of the characters: `[ ] : * ? / \`.
//! * The name cannot start or end with an apostrophe.
//! * The name shouldn't be "History" (case-insensitive) since that is reserved
//!   by Excel.
//! * The name must not be a duplicate (case-insensitive) of another worksheet
//!   name used in the workbook.
//!
//! The rules for worksheet names in Excel are explained in the [Microsoft
//! Office documentation].
//!
//! ## Setting the active worksheet
//!
//! In Excel the visible worksheet in a group of worksheets is known as the
//! active worksheet. Since only one worksheet is in the foreground at any one
//! time there can only be one active worksheet.
//!
//! With `rust_xlsxwriter` the [`Worksheet::set_active()`] method is used to
//! specify which worksheet is active. If no worksheet is set as the active
//! worksheet then the default is to have the first one active, like in Excel.
//!
//! Here is an example of making the second worksheet active:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_active.rs
//! #
//! # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     let worksheet1 = Worksheet::new();
//!     let worksheet3 = Worksheet::new();
//!     let mut worksheet2 = Worksheet::new();
//!
//!     worksheet2.set_active(true);
//!
//!     workbook.push_worksheet(worksheet1);
//!     workbook.push_worksheet(worksheet2);
//!     workbook.push_worksheet(worksheet3);
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which gives the following output:
//!
//! <img src="https://rustxlsxwriter.github.io/images/worksheet_set_active.png">
//!
//! If you have a lot of worksheets then they may not all fit on the screen at
//! the same time. In cases like that the active worksheet will still be visible
//! but its tab may not be. In those rare cases you can use the
//! [`Worksheet::set_first_tab()`] method to set the first visible tab (not
//! worksheet) in a group of worksheets.
//!
//! ## Setting worksheet tab colors
//!
//! Another way of highlighting one or more worksheets in Excel is to set the
//! tab color. With `rust_xlsxwriter` this is achieved with
//! [`Worksheet::set_tab_color()`] and a [`Color`] color:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_tab_color.rs
//! #
//! # use rust_xlsxwriter::{Color, Workbook, Worksheet, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     let mut worksheet1 = Worksheet::new();
//!     let mut worksheet2 = Worksheet::new();
//!     let mut worksheet3 = Worksheet::new();
//!     let mut worksheet4 = Worksheet::new();
//!
//!     worksheet1.set_tab_color(Color::Red);
//!     worksheet2.set_tab_color(Color::Green);
//!     worksheet3.set_tab_color(Color::RGB(0xFF9900));
//!
//!     // worksheet4 will have the default color.
//!     worksheet4.set_active(true);
//!
//!     workbook.push_worksheet(worksheet1);
//!     workbook.push_worksheet(worksheet2);
//!     workbook.push_worksheet(worksheet3);
//!     workbook.push_worksheet(worksheet4);
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which gives the following output:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_set_tab_color.png">
//!
//! ## Hiding worksheets
//!
//! Sometimes it is desirable to hide worksheets if they contain a lot of
//! intermediate data or calculations that end user doesn't need to see. With
//! `rust_xlsxwriter` this is achieved with the [`Worksheet::set_hidden()`]
//! method:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_hidden.rs
//! #
//! # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     let worksheet1 = Worksheet::new();
//!     let worksheet3 = Worksheet::new();
//!     let mut worksheet2 = Worksheet::new();
//!
//!     worksheet2.set_hidden(true);
//!
//!     workbook.push_worksheet(worksheet1);
//!     workbook.push_worksheet(worksheet2);
//!     workbook.push_worksheet(worksheet3);
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which gives the following output:
//!
//! <img src="https://rustxlsxwriter.github.io/images/worksheet_set_hidden.png">
//!
//! In Excel a hidden worksheet can not be activated or selected so
//! [`Worksheet::set_hidden()`] is mutually exclusive with the
//! [`Worksheet::set_active()`] and [`Worksheet::set_selected()`] methods. In
//! addition, since the first worksheet will default to being the active
//! worksheet, you cannot hide the first worksheet without activating another
//! sheet.
//!
//!
//! ## Selecting worksheets
//!
//! A selected worksheet has its tab highlighted. Selecting worksheets is a way
//! of grouping them together so that, for example, several worksheets could be
//! printed in one go.
//!
//! The [`Worksheet::set_selected()`] method is used to indicate that a
//! worksheet is selected in a multi-sheet workbook.
//!
//! A worksheet that has been activated via the [`Worksheet::set_active()`]
//! method will also appear as selected.
//!
//! ```
//! # // This code is available in examples/doc_worksheet_set_selected.rs
//! #
//! # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//!     let worksheet1 = Worksheet::new();
//!     let worksheet3 = Worksheet::new();
//!     let mut worksheet2 = Worksheet::new();
//!
//!     worksheet2.set_selected(true);
//!
//!     workbook.push_worksheet(worksheet1);
//!     workbook.push_worksheet(worksheet2);
//!     workbook.push_worksheet(worksheet3);
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which gives the following output. Note that Sheet 1 and Sheet2 are selected
//! but Sheet3 isn't:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_set_selected.png">
//!
//! [`Color`]: crate::Color
//! [Microsoft Office documentation]:
//!     https://support.office.com/en-ie/article/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
//!
//!
//!
//! # Grouping and outlining data
//!
//! In Excel an outline is a group of rows or columns that can be collapsed or
//! expanded to simplify hierarchical data. It is often used with the
//! `SUBTOTAL()` function. For example:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_group_rows2.png">
//!
//! Pressing the `[-]` symbol collapses and hides the grouped rows. For example
//! the following is the result of collapsing the two inner groups in the
//! previous worksheet:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_collapsed2.png">
//!
//! Outlines can be also be applied to columns:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_group_columns2.png">
//!
//! In `rust_xlsxwriter` outline groups can be created using the following
//! worksheet methods:
//!
//! - [`Worksheet::group_rows()`]
//! - [`Worksheet::group_rows_collapsed()`]
//! - [`Worksheet::group_columns()`]
//! - [`Worksheet::group_columns_collapsed()`]
//!
//! A grouping is created as follows:
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_group_rows_intro1.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #
//! #     // Add grouping over the sub-total range.
//!     worksheet.group_rows(1, 10)?;
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which creates a grouping at level 1:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro1.png">
//!
//! Hierarchical sub-groups are created by repeating the method calls for a
//! sub-range of an upper level group:
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_group_rows_intro2.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Add Level 1 grouping over the sub-total range.
//!     worksheet.group_rows(1, 10)?;
//!
//!     // Add secondary groups within the first range.
//!     worksheet.group_rows(1, 4)?;
//!     worksheet.group_rows(6, 9)?;
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! This creates the following grouping and sub-grouping at levels 1 and 2:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro2.png">
//!
//! It should be noted that Excel requires outline groups at the same level to
//! be separated by at least one row (or column) or else it will merge them into
//! a single group. This is generally to allow a subtotal row/column.
//!
//! Excel allows a maximum of 8 outline levels (including a final display
//! level). However, in practice it is rare to require more than 2 levels. All
//! available row levels are shown below.
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro3.png">
//!
//! By default Excel displays the expand and collapse `[+]` and `[-]` group
//! symbols below a row group and to the right of a column group. If required
//! the default directions can be changed using the following methods. These
//! apply to the entire worksheet:
//!
//! - [`Worksheet::group_symbols_above()`]
//! - [`Worksheet::group_symbols_to_left()`]
//!
//!
//! # Worksheet protection
//!
//! It is occasionally necessary to lock all or parts of a worksheet to prevent
//! unintentional editing. For example you may have certain fields that you want
//! a user to update but have other instruction or calculation cells that you
//! don't want modified.
//!
//! In Excel you do this by turning on the "*Review -> Sheet Protect*" option
//! and in `rust_xlsxwriter` you can use the [`Worksheet::protect()`] method:
//!
//! ```
//! # // This code is available in examples/app_worksheet_protection.rs
//! #
//! use rust_xlsxwriter::{Format, Workbook, XlsxError};
//!
//! fn main() -> Result<(), XlsxError> {
//!     // Create a new Excel file object.
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Create some format objects.
//!     let unlocked = Format::new().set_unlocked();
//!     let hidden = Format::new().set_hidden();
//!
//!     // Protect the worksheet to turn on cell locking.
//!     worksheet.protect();
//!
//!     // Examples of cell locking and hiding.
//!     worksheet.write_string(0, 0, "Cell B1 is locked. It cannot be edited.")?;
//!     worksheet.write_formula(0, 1, "=1+2")?; // Locked by default.
//!
//!     worksheet.write_string(1, 0, "Cell B2 is unlocked. It can be edited.")?;
//!     worksheet.write_formula_with_format(1, 1, "=1+2", &unlocked)?;
//!
//!     worksheet.write_string(2, 0, "Cell B3 is hidden. The formula isn't visible.")?;
//!     worksheet.write_formula_with_format(2, 1, "=1+2", &hidden)?;
//!
//!     worksheet.write_string(4, 0, "Use Menu -> Review -> Unprotect Sheet")?;
//!     worksheet.write_string(5, 0, "to remove the worksheet protection.")?;
//!
//!     worksheet.autofit();
//!
//!     // Save the file to disk.
//!     workbook.save("worksheet_protection.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! The key parts of this example are:
//!
//! - In Excel all cells have a default "locked" format so once a worksheet is
//!   protected the cells cannot be changed.
//! - To allow some cells to be edited you can set a "unlocked" format.
//! - You can also "hide" formulas in a protected worksheet.
//!
//! The output from the program will look like the following. Note that cell
//! "B2", which was unlocked in the example, has been edited.
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/app_worksheet_protection.png">
//!
//! And this is the alert you get if you try to edit a locked cell.
//!
//! <img src="https://rustxlsxwriter.github.io/images/protection_alert.png">
//!
//! ## Setting a protection password
//!
//! You can deter a user from turning off worksheet protection by adding a
//! worksheet level password using the [`Worksheet::protect_with_password()`]
//! method:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_protect_with_password.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Protect the worksheet from modification.
//!     worksheet.protect_with_password("abc123");
//! #
//! #     worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! This gives the following dialog when the user tries to unprotect the
//! worksheet.
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_password.png">
//!
//! **Note**: Worksheet level passwords in Excel offer **very weak protection**.
//! They do not encrypt your data and are very easy to deactivate. Full workbook
//! encryption is not supported by `rust_xlsxwriter`. See the section on
//! [Workbook Protection](#workbook-protection) below.
//!
//! ## Choosing which worksheet elements to protect
//!
//! Excel allows you to control which objects or actions on the worksheet that
//! are protected. The default Excel options are:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options1.png">
//!
//! Almost all the options are protected by default apart from "Select locked
//! cells" and "Select unlocked cells".
//!
//! If you wish to turn on or off any of these options you can use the
//! [`ProtectionOptions`] struct and the [`Worksheet::protect_with_options()`]
//! method. For example:
//!
//!
//! ```
//! # // This code is available in examples/doc_worksheet_protect_with_options.rs
//! #
//! # use rust_xlsxwriter::{ProtectionOptions, Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Set some of the options and use the defaults for everything else.
//!     let options = ProtectionOptions {
//!         insert_columns: true,
//!         insert_rows: true,
//!         ..ProtectionOptions::default()
//!     };
//!
//!     // Set the protection options.
//!     worksheet.protect_with_options(&options);
//! #
//! #     worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;
//! #
//! #     workbook.save("worksheet.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! This changes the allowed options to:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options2.png">
//!
//!
//! ## Workbook protection
//!
//! As noted above `rust_xlsxwriter` doesn't provide workbook level
//! encryption/protection and it is unlikely that it will be added.
//!
//! However, it is possible to encrypt an `rust_xlsxwriter` file using a third
//! party open source tool called [msoffice-crypt]. This works for macOS, Linux
//! and Windows:
//!
//! ```bash
//! msoffice-crypt.exe -e -p password clear.xlsx encrypted.xlsx
//! ```
//!
//!
//! ## Read-only workbook
//!
//! If you wish to have an Excel workbook open as read-only by default then you
//! can use the [`Workbook::read_only_recommended()`] method:
//!
//! ```
//! # // This code is available in examples/doc_workbook_read_only_recommended.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     let _worksheet = workbook.add_worksheet();
//! #
//!     workbook.read_only_recommended();
//! #
//! #     workbook.save("workbook.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! This presents the user of the file with an option to open it in "read-only"
//! mode. This means that any changes to the file can’t be saved back to the
//! same file and must be saved to a new file.
//!
//! The alert looks like this:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/workbook_read_only_recommended.png">
//!
//!
//! [msoffice-crypt]: https://github.com/herumi/msoffice
//! [`ProtectionOptions`]: crate::ProtectionOptions
//! [`Workbook::read_only_recommended()`]:
//!     crate::Workbook::read_only_recommended
//!
#![warn(missing_docs)]
mod tests;

use std::borrow::Cow;
use std::collections::btree_map::Entry;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::Cursor;
use std::io::Write;
use std::mem;
use std::sync::{Arc, Mutex, RwLock};
use std::{cmp, fmt};

#[cfg(feature = "constant_memory")]
use tempfile::tempfile_in;

#[cfg(feature = "constant_memory")]
use std::io::BufWriter;

#[cfg(feature = "constant_memory")]
use std::fs::File;

#[cfg(feature = "chrono")]
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

#[cfg(feature = "rust_decimal")]
use rust_decimal::prelude::{Decimal, ToPrimitive};

#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

#[cfg(feature = "serde")]
use crate::{
    deserialize_headers, serializer::SerializerState, CustomSerializeField,
    SerializationHeaderConfig, SerializeFieldOptions, SerializerHeader, TableData, XlsxSerialize,
};

use crate::drawing::{Drawing, DrawingCoordinates, DrawingInfo, DrawingObject, DrawingType};
use crate::error::XlsxError;
use crate::format::Format;
use crate::formula::Formula;
use crate::shared_strings_table::SharedStringsTable;
use crate::styles::Styles;
use crate::vml::VmlInfo;
use crate::xmlwriter::{
    xml_data_element, xml_data_element_only, xml_declaration, xml_empty_tag, xml_empty_tag_only,
    xml_end_tag, xml_raw_string, xml_start_tag, xml_start_tag_only, XML_WRITE_ERROR,
};
use crate::{
    utility, xmlwriter, Button, Chart, ChartEmptyCells, ChartRangeCacheData,
    ChartRangeCacheDataType, Color, ConditionalFormat, DataValidation, DataValidationErrorStyle,
    DataValidationRuleInternal, DataValidationType, ExcelDateTime, FilterCondition, FilterCriteria,
    FilterData, FilterDataType, HeaderImagePosition, HyperlinkType, Image, IntoExcelDateTime, Note,
    ObjectMovement, ProtectionOptions, Shape, Sparkline, SparklineType, Table, TableFunction, Url,
};

/// Integer type to represent a zero indexed row number. Excel's limit for rows
/// in a worksheet is 1,048,576.
pub type RowNum = u32;

/// Integer type to represent a zero indexed column number. Excel's limit for
/// columns in a worksheet is 16,384.
pub type ColNum = u16;

pub(crate) const COL_MAX: ColNum = 16_384;
pub(crate) const ROW_MAX: RowNum = 1_048_576;
pub(crate) const NUM_IMAGE_FORMATS: usize = 5;
pub(crate) const MAX_PARAMETER_LEN: usize = 255;
pub(crate) const DEFAULT_COL_WIDTH_PIXELS: u32 = 64;
pub(crate) const DEFAULT_ROW_HEIGHT_PIXELS: u32 = 20;
pub(crate) const MAX_AUTOFIT_WIDTH_PIXELS: u16 = 1790;
const DEFAULT_COL_WIDTH: f64 = 8.43;
const DEFAULT_ROW_HEIGHT: f64 = 15.0;
const MAX_STRING_LEN: usize = 32_767;
const COLUMN_LETTERS: &str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

/// The `Worksheet` struct represents an Excel worksheet. It handles operations
/// such as writing data to cells or formatting the worksheet layout.
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
pub struct Worksheet {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) name: String,
    pub(crate) sheet_index: usize,
    pub(crate) active: bool,
    pub(crate) selected: bool,
    pub(crate) visible: Visible,
    pub(crate) first_sheet: bool,
    pub(crate) has_local_string_table: bool,
    pub(crate) has_dynamic_arrays: bool,
    pub(crate) print_area_defined_name: DefinedName,
    pub(crate) repeat_row_cols_defined_name: DefinedName,
    pub(crate) autofilter_defined_name: DefinedName,
    pub(crate) autofilter_area: String,
    pub(crate) xf_formats: Vec<Format>,
    pub(crate) dxf_formats: Vec<Format>,
    pub(crate) has_vml: bool,
    pub(crate) has_hyperlink_style: bool,
    pub(crate) images: BTreeMap<(RowNum, ColNum, u32, u32), Image>,
    pub(crate) buttons_vml_info: Vec<VmlInfo>,
    pub(crate) comments_vml_info: Vec<VmlInfo>,
    pub(crate) drawing: Drawing,
    pub(crate) image_types: [bool; NUM_IMAGE_FORMATS],
    pub(crate) header_footer_images: [Option<Image>; 6],
    pub(crate) charts: BTreeMap<(RowNum, ColNum, u32, u32), Chart>,
    pub(crate) buttons: BTreeMap<(RowNum, ColNum, u32, u32), Button>,
    pub(crate) notes: BTreeMap<RowNum, BTreeMap<ColNum, Note>>,
    pub(crate) shapes: BTreeMap<(RowNum, ColNum, u32, u32), Shape>,
    pub(crate) tables: Vec<Table>,
    pub(crate) has_embedded_image_descriptions: bool,
    pub(crate) embedded_images: Vec<Image>,
    pub(crate) global_embedded_image_indices: Vec<u32>,
    pub(crate) vba_codename: Option<String>,
    pub(crate) note_authors: BTreeMap<String, usize>,
    pub(crate) vml_data_id: String,
    pub(crate) vml_shape_id: u32,
    pub(crate) is_chartsheet: bool,
    pub(crate) use_constant_memory: bool,
    pub(crate) use_inline_strings: bool,
    pub(crate) current_row: RowNum,
    pub(crate) workbook_xf_indices: Arc<RwLock<HashMap<Format, u32>>>,
    pub(crate) string_table: Arc<Mutex<SharedStringsTable>>,

    pub(crate) has_workbook_global_xfs: bool,
    pub(crate) has_workbook_global_sst: bool,
    pub(crate) background_image: Option<Image>,

    // These collections need to be reset on resave.
    drawing_rel_ids: HashMap<String, u32>,
    pub(crate) comment_relationships: Vec<(String, String, String)>,
    pub(crate) drawing_object_relationships: Vec<(String, String, String)>,
    pub(crate) drawing_relationships: Vec<(String, String, String)>,
    pub(crate) header_footer_vml_info: Vec<VmlInfo>,
    pub(crate) hyperlink_relationships: Vec<(String, String, String)>,
    pub(crate) table_relationships: Vec<(String, String, String)>,
    pub(crate) vml_drawing_relationships: Vec<(String, String, String)>,
    pub(crate) background_relationships: Vec<(String, String, String)>,

    data_table: BTreeMap<RowNum, BTreeMap<ColNum, CellType>>,
    is_writing_ahead: bool,
    merged_ranges: Vec<CellRange>,
    merged_cells: HashMap<(RowNum, ColNum), usize>,
    table_ranges: Vec<CellRange>,
    table_cells: HashMap<(RowNum, ColNum), usize>,
    dimensions: CellRange,
    xf_indices: HashMap<Format, u32>,

    dxf_indices: HashMap<Format, u32>,
    global_xf_indices: Vec<u32>,
    global_dxf_indices: Vec<u32>,
    changed_rows: HashMap<RowNum, RowOptions>,
    changed_cols: HashMap<ColNum, ColOptions>,
    page_setup_changed: bool,
    tab_color: Color,
    fit_to_page: bool,
    fit_width: u16,
    fit_height: u16,
    paper_size: u8,
    default_page_order: bool,
    right_to_left: bool,
    is_portrait: bool,
    page_view: PageView,
    zoom: u16,
    print_scale: u16,
    print_options_changed: bool,
    center_horizontally: bool,
    center_vertically: bool,
    screen_gridlines: bool,
    print_gridlines: bool,
    print_black_and_white: bool,
    print_draft: bool,
    print_headings: bool,
    header: String,
    footer: String,
    head_footer_changed: bool,
    header_footer_scale_with_doc: bool,
    header_footer_align_with_page: bool,
    margin_left: f64,
    margin_right: f64,
    margin_top: f64,
    margin_bottom: f64,
    margin_header: f64,
    margin_footer: f64,
    first_page_number: u16,
    default_result: Box<str>,
    panes: Panes,
    hyperlinks: BTreeMap<(RowNum, ColNum), Url>,
    rel_count: u32,
    protection_on: bool,
    protection_hash: u16,
    protection_options: ProtectionOptions,
    unprotected_ranges: Vec<(String, String, u16)>,
    selected_range: (String, String),
    top_left_cell: String,
    horizontal_breaks: Vec<u32>,
    vertical_breaks: Vec<u32>,
    filter_conditions: BTreeMap<ColNum, FilterCondition>,
    filter_automatic_off: bool,
    has_drawing_object_linkage: bool,
    cells_with_autofilter: HashMap<(RowNum, ColNum), (FilterType, CellRange)>,
    conditional_formats: BTreeMap<String, Vec<Box<dyn ConditionalFormat + Sync + Send>>>,
    conditional_format_order: Vec<String>,
    data_validations: BTreeMap<String, DataValidation>,
    has_conditional_formats: bool,
    use_x14_extensions: bool,
    has_x14_conditional_formats: bool,
    has_sparklines: bool,
    sparklines: Vec<Sparkline>,
    embedded_image_ids: HashMap<String, u32>,
    show_all_notes: bool,
    user_default_row_height: f64,
    hide_unused_rows: bool,
    has_sheet_data: bool,
    nan: String,
    infinity: String,
    neg_infinity: String,
    ignored_errors: HashMap<IgnoreError, String>,
    ignored_error_cells: HashSet<(RowNum, ColNum, RowNum, ColNum)>,
    max_outline_row_level: u8,
    max_outline_col_level: u8,
    outline_symbols_above: bool,
    outline_symbols_left: bool,

    #[cfg(feature = "constant_memory")]
    pub(crate) file_writer: BufWriter<File>,

    #[cfg(feature = "constant_memory")]
    write_ahead: BTreeMap<RowNum, BTreeMap<ColNum, CellType>>,

    #[cfg(feature = "serde")]
    pub(crate) serializer_state: SerializerState,
}

impl Default for Worksheet {
    fn default() -> Self {
        Self::new()
    }
}

impl Worksheet {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Worksheet object to represent an Excel worksheet.
    ///
    /// The `Worksheet::new()` constructor is used to create a new Excel
    /// worksheet object. This can be used to write data to a worksheet prior to
    /// adding it to a workbook.
    ///
    /// There are two way of creating a worksheet object with `rust_xlsxwriter`:
    /// via the [`Workbook::add_worksheet()`](crate::Workbook::add_worksheet)
    /// method and via the [`Worksheet::new()`] constructor. The first method
    /// ties the worksheet to the workbook object that will write it
    /// automatically when the file is saved, whereas the second method creates
    /// a worksheet that is independent of a workbook. This has certain
    /// advantages in keeping the worksheet free of the workbook borrow checking
    /// until you wish to add it.
    ///
    /// When working with an independent worksheet object you will need to add
    /// it to a workbook using
    /// [`Workbook::push_worksheet`](crate::Workbook::push_worksheet) in order
    /// for it to be written to a file.
    ///
    /// See also the documentation on [Creating
    /// worksheets](../worksheet/index.html#creating-worksheets) and working
    /// with the borrow checker.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating new worksheet objects and
    /// then adding them to a workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_new.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     // Create a new workbook.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Create new worksheets.
    ///     let mut worksheet1 = Worksheet::new();
    ///     let mut worksheet2 = Worksheet::new();
    ///
    ///     // Use the first workbook.
    ///     worksheet1.write_string(0, 0, "Hello")?;
    ///     worksheet1.write_string(1, 0, "Sheet1")?;
    ///
    ///     // Use the second workbook.
    ///     worksheet2.write_string(0, 0, "Hello")?;
    ///     worksheet2.write_string(1, 0, "Sheet2")?;
    ///
    ///     // Add the worksheets to the workbook.
    ///     workbook.push_worksheet(worksheet1);
    ///     workbook.push_worksheet(worksheet2);
    ///
    ///     // Save the workbook.
    ///     workbook.save("worksheets.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_new.png">
    ///
    #[allow(clippy::too_many_lines)]
    pub fn new() -> Worksheet {
        let writer = Cursor::new(Vec::with_capacity(2048));

        // Initialize the min and max dimensions with their opposite value.
        let dimensions = CellRange::default();

        let panes = Panes {
            freeze_cell: (0, 0),
            top_cell: (0, 0),
        };

        #[cfg(feature = "constant_memory")]
        let file_writer = BufWriter::new(tempfile_in(std::env::temp_dir()).unwrap());

        Worksheet {
            writer,
            name: String::new(),
            sheet_index: 0,
            active: false,
            selected: false,
            visible: Visible::Default,
            first_sheet: false,
            has_local_string_table: false,
            has_vml: false,
            has_dynamic_arrays: false,
            print_area_defined_name: DefinedName::new(),
            repeat_row_cols_defined_name: DefinedName::new(),
            autofilter_defined_name: DefinedName::new(),
            autofilter_area: String::new(),
            data_table: BTreeMap::new(),
            is_writing_ahead: false,
            dimensions,
            merged_ranges: vec![],
            merged_cells: HashMap::new(),
            tables: vec![],
            table_ranges: vec![],
            table_cells: HashMap::new(),
            xf_formats: vec![Format::default()],
            dxf_formats: vec![],
            xf_indices: HashMap::from([(Format::default(), 0)]),
            dxf_indices: HashMap::new(),
            global_xf_indices: vec![],
            global_dxf_indices: vec![],
            changed_rows: HashMap::new(),
            changed_cols: HashMap::new(),
            page_setup_changed: false,
            fit_to_page: false,
            tab_color: Color::Default,
            fit_width: 1,
            fit_height: 1,
            paper_size: 0,
            default_page_order: true,
            right_to_left: false,
            is_portrait: true,
            page_view: PageView::Normal,
            zoom: 100,
            print_scale: 100,
            print_options_changed: false,
            center_horizontally: false,
            center_vertically: false,
            screen_gridlines: true,
            print_gridlines: false,
            print_black_and_white: false,
            print_draft: false,
            print_headings: false,
            header: String::new(),
            footer: String::new(),
            head_footer_changed: false,
            header_footer_scale_with_doc: true,
            header_footer_align_with_page: true,
            margin_left: 0.7,
            margin_right: 0.7,
            margin_top: 0.75,
            margin_bottom: 0.75,
            margin_header: 0.3,
            margin_footer: 0.3,
            first_page_number: 0,
            default_result: Box::from("0"),
            panes,
            has_hyperlink_style: false,
            hyperlinks: BTreeMap::new(),
            images: BTreeMap::new(),
            shapes: BTreeMap::new(),
            drawing: Drawing::new(),
            image_types: [false; NUM_IMAGE_FORMATS],
            header_footer_images: [None, None, None, None, None, None],
            buttons_vml_info: vec![],
            comments_vml_info: vec![],
            rel_count: 0,
            protection_on: false,
            protection_hash: 0,
            protection_options: ProtectionOptions::new(),
            unprotected_ranges: vec![],
            selected_range: (String::new(), String::new()),
            top_left_cell: String::new(),
            horizontal_breaks: vec![],
            vertical_breaks: vec![],
            filter_conditions: BTreeMap::new(),
            filter_automatic_off: false,
            charts: BTreeMap::new(),
            buttons: BTreeMap::new(),
            notes: BTreeMap::new(),
            has_drawing_object_linkage: false,
            cells_with_autofilter: HashMap::new(),
            conditional_formats: BTreeMap::new(),
            conditional_format_order: vec![],
            data_validations: BTreeMap::new(),
            has_conditional_formats: false,
            use_x14_extensions: false,
            has_x14_conditional_formats: false,
            embedded_images: vec![],
            embedded_image_ids: HashMap::new(),
            global_embedded_image_indices: vec![],
            has_embedded_image_descriptions: false,
            has_sparklines: false,
            sparklines: vec![],
            vba_codename: None,
            note_authors: BTreeMap::from([("Author".to_string(), 0)]),
            show_all_notes: false,
            vml_data_id: String::new(),
            vml_shape_id: 0,
            user_default_row_height: DEFAULT_ROW_HEIGHT,
            hide_unused_rows: false,
            nan: "NAN".to_string(),
            infinity: "INF".to_string(),
            neg_infinity: "-INF".to_string(),
            ignored_errors: HashMap::new(),
            ignored_error_cells: HashSet::new(),
            max_outline_row_level: 0,
            max_outline_col_level: 0,
            outline_symbols_above: false,
            outline_symbols_left: false,
            background_image: None,

            // These collections need to be reset on resave.
            comment_relationships: vec![],
            drawing_object_relationships: vec![],
            drawing_rel_ids: HashMap::new(),
            drawing_relationships: vec![],
            header_footer_vml_info: vec![],
            hyperlink_relationships: vec![],
            table_relationships: vec![],
            vml_drawing_relationships: vec![],
            background_relationships: vec![],
            is_chartsheet: false,
            use_constant_memory: false,
            use_inline_strings: false,
            has_sheet_data: false,
            current_row: 0,
            workbook_xf_indices: Arc::new(RwLock::new(HashMap::new())),
            string_table: Arc::new(Mutex::new(SharedStringsTable::new())),
            has_workbook_global_xfs: false,
            has_workbook_global_sst: false,

            #[cfg(feature = "constant_memory")]
            file_writer,

            #[cfg(feature = "constant_memory")]
            write_ahead: BTreeMap::new(),

            #[cfg(feature = "serde")]
            serializer_state: SerializerState::new(),
        }
    }

    /// Create a new Worksheet object to represent an Excel chartsheet
    ///
    /// The `Worksheet::new_chartsheet()` constructor is used to create a new
    /// Excel "chartsheet" object. A Chartsheet in Excel is a specialized type
    /// of worksheet that doesn't have cells but instead is used to display a
    /// single chart. It supports worksheet display options such as headers and
    /// footers, margins, tab selection and print properties.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chartsheet.png">
    ///
    /// The `Worksheet::new_chartsheet()` method returns a standard
    /// [`Worksheet`] that can be used as a chartsheet by adding a chart using
    /// the [`Worksheet::insert_image()`] method. It supports most other
    /// worksheet functions apart from those that  are cell based such as
    /// `Worksheet::write()` and the other `Worksheet::insert*()` methods.
    ///
    /// # Examples
    ///
    /// A simple chartsheet example that creates a workbook with a worksheet and
    /// a chartsheet as shown in the image above.
    ///
    /// ```
    /// # // This code is available in examples/doc_chartsheet.rs
    /// #
    /// use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some data for the chart.
    ///     worksheet.write(0, 0, 10)?;
    ///     worksheet.write(1, 0, 60)?;
    ///     worksheet.write(2, 0, 30)?;
    ///     worksheet.write(3, 0, 10)?;
    ///     worksheet.write(4, 0, 50)?;
    ///
    ///     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$5");
    ///
    ///     // Create a new chartsheet.
    ///     let chartsheet = workbook.add_chartsheet();
    ///
    ///     // Add the chart to the chartsheet.
    ///     chartsheet.insert_chart(0, 0, &chart)?;
    ///
    ///     // Save the file.
    ///     workbook.save("chart.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    pub fn new_chartsheet() -> Worksheet {
        let protection_options = ProtectionOptions {
            edit_objects: true,
            ..Default::default()
        };

        Worksheet {
            is_chartsheet: true,
            is_portrait: false,
            protection_options,
            ..Default::default()
        }
    }

    /// Set the worksheet name.
    ///
    /// Set the worksheet name. If no name is set the default Excel convention
    /// will be followed (Sheet1, Sheet2, etc.) in the order the worksheets are
    /// created.
    ///
    /// # Parameters
    ///
    /// - `name`: The worksheet name. It must follow the Excel rules, shown
    ///   below.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name cannot be
    ///   blank.
    /// - [`XlsxError::SheetnameLengthExceeded`] - Worksheet name exceeds
    ///   Excel's limit of 31 characters.
    /// - [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name
    ///   cannot contain invalid characters: `[ ] : * ? / \`
    /// - [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name
    ///   cannot start or end with an apostrophe.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting user defined worksheet names
    /// and the default values when a name isn't set.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let _worksheet1 = workbook.add_worksheet(); // Defaults to Sheet1
    ///     let _worksheet2 = workbook.add_worksheet().set_name("Foglio2");
    ///     let _worksheet3 = workbook.add_worksheet().set_name("Data");
    ///     let _worksheet4 = workbook.add_worksheet(); // Defaults to Sheet4
    ///
    /// #     workbook.save("worksheets.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_name.png">
    ///
    /// The worksheet name must be a valid Excel worksheet name, i.e:
    ///
    /// - The name is less than 32 characters.
    /// - The name isn't blank.
    /// - The name doesn't contain any of the characters: `[ ] : * ? / \`.
    /// - The name doesn't start or end with an apostrophe.
    /// - The name shouldn't be "History" (case-insensitive) since that is
    ///   reserved by Excel.
    /// - It must not be a duplicate of another worksheet name used in the
    ///   workbook.
    ///
    /// The rules for worksheet names in Excel are explained in the [Microsoft
    /// Office documentation].
    ///
    /// [Microsoft Office documentation]:
    ///     https://support.office.com/en-ie/article/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
    ///
    pub fn set_name(&mut self, name: impl Into<String>) -> Result<&mut Worksheet, XlsxError> {
        let name = name.into();

        let error_message = format!(r#"worksheet.set_name("{name}")"#);
        utility::validate_sheetname(&name, &error_message)?;

        self.name = name;

        Ok(self)
    }

    /// Get the worksheet name.
    ///
    /// Get the worksheet name that was set automatically such as Sheet1,
    /// Sheet2, etc., or that was set by the user using
    /// [`Worksheet::set_name()`].
    ///
    /// The worksheet name can be used to get a reference to a worksheet object
    /// using the
    /// [`Workbook::worksheet_from_name()`](crate::Workbook::worksheet_from_name)
    /// method.
    ///
    /// # Examples
    ///
    /// The following example demonstrates getting a worksheet name.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_name.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Try name() using a default sheet name.
    ///     let worksheet = workbook.add_worksheet();
    ///     assert_eq!("Sheet1", worksheet.name());
    ///
    ///     // Try name() using a user defined sheet name.
    ///     let worksheet = workbook.add_worksheet().set_name("Data")?;
    ///     assert_eq!("Data", worksheet.name());
    /// #
    /// #    workbook.save("workbook.xlsx")?;
    /// #
    /// #    Ok(())
    /// # }
    /// ```
    ///
    pub fn name(&self) -> String {
        self.name.clone()
    }

    /// Write generic data to a cell.
    ///
    /// The `write()` method writes data that implements [`IntoExcelData`] to a
    /// worksheet cell.
    ///
    /// The types currently supported are:
    /// - String types: [`&str`], [`String`], `&String` and `Cow<'_, str>`.
    /// - Numbers that convert [`Into`] [`f64`]. Also, u64 and i64 are supported
    ///   with loss of precision outside Excel's integer range of +/-
    ///   999,999,999,999,999 (15 digits).
    /// - [`bool`]
    /// - [`ExcelDateTime`].
    /// - [`Formula`].
    /// - [`Url`].
    /// - [`Option<T>`]: If `T` is a supported type then the [`Some<T>`] value
    ///   is written. The [`None`] value is ignored.
    /// - [`Result<T, E>`]: If `T` and `E` are supported types then the `T` or
    ///   `E` value is written depending on the result.
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - [`chrono::NaiveDateTime`].
    /// - [`chrono::NaiveDate`].
    /// - [`chrono::NaiveTime`].
    ///
    /// [`Chrono`]: https://docs.rs/chrono/latest/chrono/index.html
    /// [`chrono::NaiveDate`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    /// [`chrono::NaiveTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    /// [`chrono::NaiveDateTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// If the `rust_decimal` feature is enabled you can write the [`Decimal`] type via [`rust_decimal`].
    /// This requires that the `Decimal` can be represented as a `f64` in Excel.
    ///
    /// [`Decimal`]: https://docs.rs/rust_decimal/latest/rust_decimal/struct.Decimal.html
    /// [`rust_decimal`]: https://docs.rs/rust_decimal/latest/rust_decimal
    ///
    /// Users can also use this method to write their own data types to Excel by
    /// implementing the [`IntoExcelData`] trait.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: An type that implements the  [`IntoExcelData`] trait.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    pub fn write(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: impl IntoExcelData,
    ) -> Result<&mut Worksheet, XlsxError> {
        data.write(self, row, col)
    }

    /// Write formatted generic data to a cell.
    ///
    /// The `write_with_format()` method writes formatted data that implements
    /// [`IntoExcelData`] to a worksheet cell.
    ///
    /// The types currently supported are:
    /// - String types: [`&str`], [`String`], `&String` and `Cow<'_, str>`.
    /// - Numbers that convert [`Into`] [`f64`]. Also, u64 and i64 are supported
    ///   with loss of precision outside Excel's integer range of +/-
    ///   999,999,999,999,999 (15 digits).
    /// - [`bool`]
    /// - [`ExcelDateTime`].
    /// - [`Formula`].
    /// - [`Url`].
    /// - [`Option<T>`]: If `T` is a supported type then [`Some<T>`] is written
    ///   as a formatted value or [`None`] is written as a formatted blank cell.
    /// - [`Result<T, E>`]: If `T` and `E` are supported types then the `T` or
    ///   `E` value is written depending on the result.
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - [`chrono::NaiveDateTime`].
    /// - [`chrono::NaiveDate`].
    /// - [`chrono::NaiveTime`].
    ///
    /// [`Chrono`]: https://docs.rs/chrono/latest/chrono/index.html
    /// [`chrono::NaiveDate`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    /// [`chrono::NaiveTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    /// [`chrono::NaiveDateTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// Users can also use this method to write their own data types to Excel by
    /// implementing the [`IntoExcelData`] trait.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: An type that implements the  [`IntoExcelData`] trait.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    pub fn write_with_format<'a, T>(
        &'a mut self,
        row: RowNum,
        col: ColNum,
        data: T,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError>
    where
        T: IntoExcelData,
    {
        data.write_with_format(self, row, col, format)
    }

    /// Write an array like data structure as a row of data to a worksheet.
    ///
    /// Write an array of data horizontally rightwards starting from the initial
    /// `row, col` cell.
    ///
    /// This methods works for arrays or array-like data structures that
    /// implement [`IntoIterator`] and that contain a data type that implements
    /// [`IntoExcelData`].
    ///
    /// See also [`Worksheet::write_column()`] for a similar function that works
    /// in an orthogonal direction.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: Arrays or array-like data structures that implement
    ///   [`IntoIterator`] and that contain a data type that implements
    ///   [`IntoExcelData`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array of data as a row to
    /// a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_row.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some array data to write.
    ///     let data = [1, 2, 3, 4, 5];
    ///
    ///     // Write the array data as a row.
    ///     worksheet.write_row(0, 0, data)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_row.png">
    ///
    /// An example of writing arrays of data using the `rust_xlsxwriter`
    /// library. Array in this context means Rust arrays or arrays like data
    /// types that implement [`IntoIterator`]. The array must also contain data
    /// types that implement `rust_xlsxwriter`'s [`IntoExcelData`].
    ///
    /// ```
    /// # // This code is available in examples/app_write_arrays.rs
    /// #
    /// use rust_xlsxwriter::{Format, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a format for the headings.
    ///     let heading = Format::new().set_bold().set_font_color("#0000CC");
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some array data to write.
    ///     let numbers = [1, 2, 3, 4, 5];
    ///     let words = ["Hello"; 5];
    ///     let matrix = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///
    ///     // Write the array data as columns.
    ///     worksheet.write_with_format(0, 0, "Column data", &heading)?;
    ///     worksheet.write_column(1, 0, numbers)?;
    ///     worksheet.write_column(1, 1, words)?;
    ///
    ///     // Write the array data as rows.
    ///     worksheet.write_with_format(0, 4, "Row data", &heading)?;
    ///     worksheet.write_row(1, 4, numbers)?;
    ///     worksheet.write_row(2, 4, words)?;
    ///
    ///     // Write the matrix data as an array or rows and as an array of columns.
    ///     worksheet.write_with_format(7, 4, "Row matrix", &heading)?;
    ///     worksheet.write_row_matrix(8, 4, matrix)?;
    ///
    ///     worksheet.write_with_format(7, 0, "Column matrix", &heading)?;
    ///     worksheet.write_column_matrix(8, 0, matrix)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("arrays.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/app_write_arrays.png">
    ///
    pub fn write_row<I>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: I,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        I: IntoIterator,
        I::Item: IntoExcelData,
    {
        let mut col = col;
        for item in data {
            self.write(row, col, item)?;
            col += 1;
        }

        Ok(self)
    }

    /// Write an array like data structure as a row of data to a worksheet, with
    /// formatting.
    ///
    /// This method is similar to [`Worksheet::write_row()`] except you can also
    /// specify a format for the data.
    ///
    /// See [`Worksheet::write_row()`] above for details.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: Arrays or array-like data structures that implement
    ///   [`IntoIterator`] and that contain a data type that implements
    ///   [`IntoExcelData`].
    /// - `format`: The [`Format`] property for the data.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    pub fn write_row_with_format<I>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: I,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        I: IntoIterator,
        I::Item: IntoExcelData,
    {
        let mut col = col;
        for item in data {
            self.write_with_format(row, col, item, format)?;
            col += 1;
        }

        Ok(self)
    }

    /// Write an array like data structure as a column of data to a worksheet.
    ///
    /// Write an array of data vertically downwards starting from the initial
    /// `row, col` cell.
    ///
    /// This methods works for arrays or array-like data structures that
    /// implement [`IntoIterator`] and that contain a data type that implements
    /// [`IntoExcelData`].
    ///
    /// See also [`Worksheet::write_row()`] for a similar function that works in
    /// an orthogonal direction.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: Arrays or array-like data structures that implement
    ///   [`IntoIterator`] and that contain a data type that implements
    ///   [`IntoExcelData`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array of data as a column
    /// to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_column.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some array data to write.
    ///     let data = [1, 2, 3, 4, 5];
    ///
    ///     // Write the array data as a column.
    ///     worksheet.write_column(0, 0, data)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_column.png">
    ///
    pub fn write_column<I>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: I,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        I: IntoIterator,
        I::Item: IntoExcelData,
    {
        let mut row = row;
        for item in data {
            self.write(row, col, item)?;
            row += 1;
        }

        Ok(self)
    }

    /// Write an array like data structure as a column of data to a worksheet, with
    /// formatting.
    ///
    /// This method is similar to [`Worksheet::write_column()`] except you can also
    /// specify a format for the data.
    ///
    /// See [`Worksheet::write_column()`] above for details.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: Arrays or array-like data structures that implement
    ///   [`IntoIterator`] and that contain a data type that implements
    ///   [`IntoExcelData`].
    /// - `format`: The [`Format`] property for the data.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    pub fn write_column_with_format<I>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: I,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        I: IntoIterator,
        I::Item: IntoExcelData,
    {
        let mut row = row;
        for item in data {
            self.write_with_format(row, col, item, format)?;
            row += 1;
        }

        Ok(self)
    }

    /// Write an array of row arrays to a worksheet.
    ///
    /// Write an array of row arrays vertically downwards starting from the
    /// initial `row, col` cell.
    ///
    /// This methods works for 2D arrays or array-like data structures that
    /// implement [`IntoIterator`] and that contain a data type that implements
    /// [`IntoExcelData`].
    ///
    /// See also [`Worksheet::write_column_matrix()`] for a similar function
    /// that works in an orthogonal direction.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: 2D arrays or array-like data structures that implement
    ///   [`IntoIterator`] and that contain a data type that implements
    ///   [`IntoExcelData`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array of row arrays to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_row_matrix.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some array data to write.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///
    ///     // Write the array data as a series of rows.
    ///     worksheet.write_row_matrix(0, 0, data)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_row_matrix.png">
    ///
    pub fn write_row_matrix<I, II>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: I,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        I: IntoIterator,
        I::Item: IntoIterator<Item = II>,
        II: IntoExcelData,
    {
        let mut row = row;
        for item in data {
            self.write_row(row, col, item)?;
            row += 1;
        }

        Ok(self)
    }

    /// Write an array of column arrays to a worksheet.
    ///
    /// Write an array of column arrays horizontally rightwards starting from
    /// the initial `row, col` cell.
    ///
    /// This methods works for 2D arrays or array-like data structures that
    /// implement [`IntoIterator`] and that contain a data type that implements
    /// [`IntoExcelData`].
    ///
    /// See also [`Worksheet::write_row_matrix()`] for a similar function that
    /// works in an orthogonal direction.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data`: 2D arrays or array-like data structures that implement
    ///   [`IntoIterator`] and that contain a data type that implements
    ///   [`IntoExcelData`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array of column arrays to
    /// a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_column_matrix.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some array data to write.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///
    ///     // Write the array data as a series of columns.
    ///     worksheet.write_column_matrix(0, 0, data)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_column_matrix.png">
    ///
    pub fn write_column_matrix<I, II>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: I,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        I: IntoIterator,
        I::Item: IntoIterator<Item = II>,
        II: IntoExcelData,
    {
        let mut col = col;
        for item in data {
            self.write_column(row, col, item)?;
            col += 1;
        }

        Ok(self)
    }

    /// Write an unformatted number to a cell.
    ///
    /// Write an unformatted number to a worksheet cell. To write a formatted
    /// number see the [`Worksheet::write_number_with_format()`] method below.
    ///
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's [`f64`] type. This method will accept any rust
    /// type that will convert [`Into`] a f64. These include i8, u8, i16, u16,
    /// i32, u32 and f32 but not i64 or u64, see below.
    ///
    /// IEEE 754 Doubles and f64 have around 15 digits of precision. Anything
    /// beyond that cannot be stored as a number by Excel without a loss of
    /// precision and may need to be stored as a string instead.
    ///
    /// [IEEE 754]: https://en.wikipedia.org/wiki/IEEE_754
    ///
    /// For i64/u64 you can cast the numbers `as f64` which will allow you to
    /// store the number with a loss of precision outside Excel's integer range
    /// of +/- 999,999,999,999,999 (15 digits).
    ///
    /// Excel doesn't have handling for NaN or INF floating point numbers. These
    /// will be stored as the strings "NAN", "INF", and "-INF" strings or the
    /// values set with [`Worksheet::set_nan_value()`],
    /// [`Worksheet::set_infinity_value()`] or
    /// [`Worksheet::set_neg_infinity_value()`].
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `number`: The number to write to the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing unformatted numbers to an
    /// Excel worksheet. Any numeric type that will convert [`Into`] f64 can be
    /// transferred to Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_number.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some different Rust number types to a worksheet.
    ///     worksheet.write_number(0, 0, 1_u8)?;
    ///     worksheet.write_number(1, 0, 2_i16)?;
    ///     worksheet.write_number(2, 0, 3_u32)?;
    ///     worksheet.write_number(3, 0, 4_f32)?;
    ///     worksheet.write_number(4, 0, 5_f64)?;
    ///
    ///     // Write some numbers with implicit types.
    ///     worksheet.write_number(5, 0, 1234)?;
    ///     worksheet.write_number(6, 0, 1234.5)?;
    ///
    ///     // Note Excel normally ignores trailing decimal zeros
    ///     // when the number is unformatted.
    ///     worksheet.write_number(7, 0, 1234.50000)?;
    /// #
    /// #     workbook.save("numbers.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_number.png">
    ///
    pub fn write_number(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: impl Into<f64>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_number(row, col, number, None)
    }

    /// Write a formatted number to a worksheet cell.
    ///
    /// Write a number with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's [`f64`] type. This method will accept any rust
    /// type that will convert [`Into`] a f64. These include i8, u8, i16, u16,
    /// i32, u32 and f32 but not i64 or u64, see below.
    ///
    /// IEEE 754 Doubles and f64 have around 15 digits of precision. Anything
    /// beyond that cannot be stored as a number by Excel without a loss of
    /// precision and may need to be stored as a string instead.
    ///
    /// [IEEE 754]: https://en.wikipedia.org/wiki/IEEE_754
    ///
    /// For i64/u64 you can cast the numbers `as f64` which will allow you to
    /// store the number with a loss of precision outside Excel's integer range
    /// of +/- 999,999,999,999,999 (15 digits).
    ///
    /// Excel doesn't have handling for NaN or INF floating point numbers. These
    /// will be stored as the strings "NAN", "INF", and "-INF" strings or the
    /// values set with [`Worksheet::set_nan_value()`],
    /// [`Worksheet::set_infinity_value()`] or
    /// [`Worksheet::set_neg_infinity_value()`].
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `number`: The number to write to the cell.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting different formatting for
    /// numbers in an Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_number_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Create some formats to use with the numbers below.
    ///     let number_format = Format::new().set_num_format("#,##0.00");
    ///     let currency_format = Format::new().set_num_format("€#,##0.00");
    ///     let percentage_format = Format::new().set_num_format("0.0%");
    ///     let bold_italic_format = Format::new().set_bold().set_italic();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_number_with_format(0, 0, 1234.5, &number_format)?;
    ///     worksheet.write_number_with_format(1, 0, 1234.5, &currency_format)?;
    ///     worksheet.write_number_with_format(2, 0, 0.3300, &percentage_format)?;
    ///     worksheet.write_number_with_format(3, 0, 1234.5, &bold_italic_format)?;
    /// #
    /// #     workbook.save("numbers.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_number_with_format.png">
    ///
    ///
    pub fn write_number_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: impl Into<f64>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_number(row, col, number.into(), Some(format))
    }

    /// Write an unformatted string to a worksheet cell.
    ///
    /// Write an unformatted string to a worksheet cell. To write a formatted
    /// string see the [`Worksheet::write_string_with_format()`] method below.
    ///
    /// Excel only supports UTF-8 text in the xlsx file format. Any Rust UTF-8
    /// encoded string can be written with this method. The maximum string size
    /// supported by Excel is 32,767 characters.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `string`: The string to write to the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing some strings to a worksheet.
    /// The UTF-8 strings are taken from the UTF-8 example in the [Rust
    /// Programming Language] book.
    ///
    /// [Rust Programming Language]:
    ///     https://doc.rust-lang.org/book/ch08-02-strings.html#creating-a-new-string
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_string.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #   // Create a new Excel file object.
    /// #   let mut workbook = Workbook::new();
    /// #
    /// #   // Add a worksheet to the workbook.
    /// #   let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write some strings to the worksheet.
    ///     worksheet.write_string(0,  0, "السلام عليكم")?;
    ///     worksheet.write_string(1,  0, "Dobrý den")?;
    ///     worksheet.write_string(2,  0, "Hello")?;
    ///     worksheet.write_string(3,  0, "שָׁלוֹם")?;
    ///     worksheet.write_string(4,  0, "नमस्ते")?;
    ///     worksheet.write_string(5,  0, "こんにちは")?;
    ///     worksheet.write_string(6,  0, "안녕하세요")?;
    ///     worksheet.write_string(7,  0, "你好")?;
    ///     worksheet.write_string(8,  0, "Olá")?;
    ///     worksheet.write_string(9,  0, "Здравствуйте")?;
    ///     worksheet.write_string(10, 0, "Hola")?;
    /// #
    /// #     workbook.save("strings.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_string.png">
    ///
    pub fn write_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: impl Into<String>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string.into(), None)
    }

    /// Write a formatted string to a worksheet cell.
    ///
    /// Write a string with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// Excel only supports UTF-8 text in the xlsx file format. Any Rust UTF-8
    /// encoded string can be written with this method. The maximum string
    /// size supported by Excel is 32,767 characters.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `string`: The string to write to the cell.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting different formatting for
    /// numbers in an Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_string_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Create some formats to use in the worksheet.
    ///     let bold_format = Format::new().set_bold();
    ///     let italic_format = Format::new().set_italic();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some strings with formatting.
    ///     worksheet.write_string_with_format(0, 0, "Hello",     &bold_format)?;
    ///     worksheet.write_string_with_format(1, 0, "שָׁלוֹם",      &bold_format)?;
    ///     worksheet.write_string_with_format(2, 0, "नमस्ते",      &italic_format)?;
    ///     worksheet.write_string_with_format(3, 0, "こんにちは", &italic_format)?;
    /// #
    /// #     workbook.save("strings.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_string_with_format.png">
    ///
    pub fn write_string_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: impl Into<String>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string.into(), Some(format))
    }

    /// Write a "rich" string with multiple formats to a worksheet cell.
    ///
    /// The `write_rich_string()` method is used to write strings with multiple
    /// font formats within the string. For example strings like "This is
    /// **bold** and this is *italic*". For strings with a single format you can
    /// use the more common [`Worksheet::write_string_with_format()`] method.
    ///
    /// The basic rule is to break the string into pairs of [`Format`] and
    /// [`str`] fragments. So if we look at the above string again:
    ///
    /// &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This is **bold** and this is
    /// *italic*
    ///
    /// The would be broken down into 4 fragments:
    ///
    /// ```text
    ///      default: |This is |
    ///      bold:    |bold|
    ///      default: | and this is |
    ///      italic:  |italic|
    /// ```
    ///
    /// This should then be converted to an array of [`Format`] and [`str`]
    /// tuples:
    ///
    /// ```text
    ///     let segments = [
    ///        (&default, "This is "),
    ///        (&red,     "red"),
    ///        (&default, " and this is "),
    ///        (&blue,    "blue"),
    ///     ];
    /// ```
    ///
    /// See the full example below.
    ///
    /// For the default format segments you can use [`Format::default()`].
    ///
    /// Note, only the Font elements of the [`Format`] are used by Excel in rich
    /// strings. For example it isn't possible in Excel to highlight part of the
    /// string with a yellow background. It is possible to have a yellow
    /// background for the entire cell or to format other cell properties using
    /// an additional [`Format`] object and the
    /// [`Worksheet::write_rich_string_with_format()`] method, see below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `rich_string`: An array reference of `(&Format, &str)` tuples. See the
    ///   Errors section below for the restrictions.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::ParameterError`] - The following error cases will raise a
    ///   `ParameterError` error:
    ///   * If any of the str elements is empty. Excel doesn't allow this.
    ///   * If there isn't at least one `(&Format, &str)` tuple element in the
    ///     `rich_string` parameter array. Strictly speaking there should be at
    ///     least 2 tuples to make a rich string, otherwise it is just a normal
    ///     formatted string. However, Excel allows it.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a "rich" string with multiple
    /// formats.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_rich_string.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.set_column_width(0, 30)?;
    /// #
    ///     // Add some formats to use in the rich strings.
    ///     let default = Format::default();
    ///     let red = Format::new().set_font_color(Color::Red);
    ///     let blue = Format::new().set_font_color(Color::Blue);
    ///
    ///     // Write a Rich strings with multiple formats.
    ///     let segments = [
    ///         (&default, "This is "),
    ///         (&red,     "red"),
    ///         (&default, " and this is "),
    ///         (&blue,    "blue"),
    ///     ];
    ///     worksheet.write_rich_string(0, 0, &segments)?;
    ///
    ///     // It is possible, and idiomatic, to use slices as the string segments.
    ///     let text = "This is blue and this is red";
    ///     let segments = [
    ///         (&default, &text[..8]),
    ///         (&blue,    &text[8..12]),
    ///         (&default, &text[12..25]),
    ///         (&red,     &text[25..]),
    ///     ];
    ///     worksheet.write_rich_string(1, 0, &segments)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_rich_string.png">
    ///
    pub fn write_rich_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        rich_string: &[(&Format, &str)],
    ) -> Result<&mut Worksheet, XlsxError> {
        let (string, raw_string) = Self::get_rich_string(rich_string)?;

        self.store_rich_string(row, col, &string, &raw_string, None)
    }

    /// Write a "rich" string with multiple formats to a worksheet cell, with an
    /// additional cell format.
    ///
    /// The `write_rich_string_with_format()` method is used to write strings
    /// with multiple font formats within the string. For example strings like
    /// "This is **bold** and this is *italic*". It also allows you to add an
    /// additional [`Format`] to the cell so that you can, for example, center
    /// the text in the cell.
    ///
    /// The syntax for creating and using `(&Format, &str)` tuples to create the
    /// rich string is shown above in [`Worksheet::write_rich_string()`].
    ///
    /// For strings with a single format you can use the more common
    /// [`Worksheet::write_string_with_format()`] method.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `rich_string`: An array reference of `(&Format, &str)` tuples. See the
    ///   Errors section below for the restrictions.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::ParameterError`] - The following error cases will raise a
    ///   `ParameterError` error:
    ///   * If any of the str elements is empty. Excel doesn't allow this.
    ///   * If there isn't at least one `(&Format, &str)` tuple element in the
    ///     `rich_string` parameter array. Strictly speaking there should be at
    ///     least 2 tuples to make a rich string, otherwise it is just a normal
    ///     formatted string. However, Excel allows it.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a "rich" string with multiple
    /// formats, and an additional cell format.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_rich_string_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatAlign, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.set_column_width(0, 30)?;
    /// #
    ///     // Add some formats to use in the rich strings.
    ///     let default = Format::default();
    ///     let red = Format::new().set_font_color(Color::Red);
    ///     let blue = Format::new().set_font_color(Color::Blue);
    ///
    ///     // Write a rich strings with multiple formats.
    ///     let segments = [
    ///         (&default, "This is "),
    ///         (&red,     "red"),
    ///         (&default, " and this is "),
    ///         (&blue,    "blue"),
    ///     ];
    ///     worksheet.write_rich_string(0, 0, &segments)?;
    ///
    ///     // Add an extra format to use for the entire cell.
    ///     let center = Format::new().set_align(FormatAlign::Center);
    ///
    ///     // Write the rich string again with the cell format.
    ///     worksheet.write_rich_string_with_format(2, 0, &segments, &center)?;
    ///
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_rich_string_with_format.png">
    ///
    pub fn write_rich_string_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        rich_string: &[(&Format, &str)],
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let (string, raw_string) = Self::get_rich_string(rich_string)?;

        self.store_rich_string(row, col, &string, &raw_string, Some(format))
    }

    /// Write an unformatted formula to a worksheet cell.
    ///
    /// Write an unformatted Excel formula to a worksheet cell. See also the
    /// documentation on working with formulas at [`Formula`].
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or [`Formula`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formulas with formatting to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some formulas to the worksheet.
    ///     worksheet.write_formula(0, 0, "=B3 + B4")?;
    ///     worksheet.write_formula(1, 0, "=SIN(PI()/4)")?;
    ///     worksheet.write_formula(2, 0, "=SUM(B1:B5)")?;
    ///     worksheet.write_formula(3, 0, r#"=IF(A3>1,"Yes", "No")"#)?;
    ///     worksheet.write_formula(4, 0, "=AVERAGE(1, 2, 3, 4)")?;
    ///     worksheet.write_formula(5, 0, r#"=DATEVALUE("1-Jan-2023")"#)?;
    /// #
    /// #     workbook.save("formulas.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_formula.png">
    ///
    pub fn write_formula(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: impl Into<Formula>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_formula(row, col, formula.into(), None)
    }

    /// Write a formatted formula to a worksheet cell.
    ///
    /// Write a formula with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// See also the documentation on working with formulas at [`Formula`].
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or [`Formula`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formulas with formatting to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_formula_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Create some formats to use in the worksheet.
    ///     let bold_format = Format::new().set_bold();
    ///     let italic_format = Format::new().set_italic();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write some formulas with formatting.
    ///     worksheet.write_formula_with_format(0, 0, "=1+2+3", &bold_format)?;
    ///     worksheet.write_formula_with_format(1, 0, "=A1*2", &bold_format)?;
    ///     worksheet.write_formula_with_format(2, 0, "=SIN(PI()/4)", &italic_format)?;
    ///     worksheet.write_formula_with_format(3, 0, "=AVERAGE(1, 2, 3, 4)", &italic_format)?;
    /// #
    /// #     workbook.save("formulas.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_formula_with_format.png">
    ///
    pub fn write_formula_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: impl Into<Formula>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_formula(row, col, formula.into(), Some(format))
    }

    /// Write an  array formula to a worksheet cell.
    ///
    /// The `write_array_formula()` method writes an array formula to a
    /// cell range. In Excel an array formula is a formula that performs a
    /// calculation on a range of values. It can return a single value or a
    /// range/"array" of values.
    ///
    /// An array formula is displayed with a pair of curly brackets around the
    /// formula like this: `{=SUM(A1:B1*A2:B2)}`. The `write_array()`
    /// method doesn't require actually require these so you can omit them in
    /// the formula, and the equal sign, if you wish like this:
    /// `SUM(A1:B1*A2:B2)`.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or [`Formula`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row or column is larger
    ///   than the last row or column.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array formulas to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_array_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    /// #    // Write some test data.
    /// #    worksheet.write_number(0, 1, 500)?;
    /// #    worksheet.write_number(0, 2, 300)?;
    /// #    worksheet.write_number(1, 1, 10)?;
    /// #    worksheet.write_number(1, 2, 15)?;
    /// #
    ///     // Write an array formula that returns a single value.
    ///     worksheet.write_array_formula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}")?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_array_formula.png">
    ///
    pub fn write_array_formula(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: impl Into<Formula>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            formula.into(),
            None,
            false,
        )
    }

    /// Write a formatted array formula to a worksheet cell.
    ///
    /// Write an array formula with formatting to a worksheet cell. The format
    /// is set via a [`Format`] struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// The `write_array()` method writes an array formula to a cell
    /// range. In Excel an array formula is a formula that performs a
    /// calculation on a range of values. It can return a single value or a
    /// range/"array" of values.
    ///
    /// An array formula is displayed with a pair of curly brackets around the
    /// formula like this: `{=SUM(A1:B1*A2:B2)}`. The `write_array()`
    /// method doesn't require actually require these so you can omit them in
    /// the formula, and the equal sign, if you wish like this:
    /// `SUM(A1:B1*A2:B2)`.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or [`Formula`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing an array formula with
    /// formatting to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_array_formula_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    /// #    // Add a format.
    /// #    let bold = Format::new().set_bold();
    /// #
    /// #    // Write some test data.
    /// #    worksheet.write_number(0, 1, 500)?;
    /// #    worksheet.write_number(0, 2, 300)?;
    /// #    worksheet.write_number(1, 1, 10)?;
    /// #    worksheet.write_number(1, 2, 15)?;
    /// #
    ///     // Write an array formula that returns a single value.
    ///     worksheet.write_array_formula_with_format(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", &bold)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_array_formula_with_format.png">
    ///
    pub fn write_array_formula_with_format(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: impl Into<Formula>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            formula.into(),
            Some(format),
            false,
        )
    }

    /// Write a dynamic array formula to a worksheet cell or range of cells.
    ///
    /// The `write_dynamic_array_formula()` function writes an Excel 365
    /// dynamic array formula to a cell range. Some examples of functions that
    /// return dynamic arrays are:
    ///
    /// - `FILTER()`
    /// - `RANDARRAY()`
    /// - `SEQUENCE()`
    /// - `SORTBY()`
    /// - `SORT()`
    /// - `UNIQUE()`
    /// - `XLOOKUP()`
    /// - `XMATCH()`
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or [`Formula`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates a static function which generally
    /// returns one value turned into a dynamic array function which returns a
    /// range of values.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_dynamic_array_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Write a dynamic formula using a static function.
    ///     worksheet.write_dynamic_array_formula(0, 1, 0, 1, "=LEN(A1:A3)")?;
    /// #
    /// #     // Write some data for the function to operate on.
    /// #     worksheet.write_string(0, 0, "Foo")?;
    /// #     worksheet.write_string(1, 0, "Food")?;
    /// #     worksheet.write_string(2, 0, "Frood")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_dynamic_array_formula.png">
    ///
    pub fn write_dynamic_array_formula(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: impl Into<Formula>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            formula.into(),
            None,
            true,
        )
    }

    /// Write a formatted dynamic array formula to a worksheet cell or range of
    /// cells.
    ///
    /// The `write_dynamic_array_formula_with_format()` function writes an Excel
    /// 365 dynamic array formula to a cell range. Some examples of functions
    /// that return dynamic arrays are:
    ///
    /// - `FILTER()`
    /// - `RANDARRAY()`
    /// - `SEQUENCE()`
    /// - `SORTBY()`
    /// - `SORT()`
    /// - `UNIQUE()`
    /// - `XLOOKUP()`
    /// - `XMATCH()`
    ///
    /// The format is set via a [`Format`] struct which can control the font or
    /// color or properties such as bold and italic.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below or use the [`Worksheet::write_dynamic_formula_with_format()`]
    /// method.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or
    ///   [`Formula`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row or column is larger
    ///   than the last row or column.
    ///
    /// # Examples
    ///
    /// The following example demonstrates a static function which generally
    /// returns one value turned into a dynamic array function which returns a
    /// range of values.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_dynamic_array_formula_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Write a dynamic formula using a static function.
    ///     worksheet.write_dynamic_array_formula_with_format(0, 1, 0, 1, "=LEN(A1:A3)", &bold)?;
    /// #
    /// #     // Write some data for the function to operate on.
    /// #     worksheet.write_string(0, 0, "Foo")?;
    /// #     worksheet.write_string(1, 0, "Food")?;
    /// #     worksheet.write_string(2, 0, "Frood")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_dynamic_array_formula_with_format.png">
    ///
    pub fn write_dynamic_array_formula_with_format(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: impl Into<Formula>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            formula.into(),
            Some(format),
            true,
        )
    }

    /// Write a dynamic formula to a worksheet cell.
    ///
    /// The `write_dynamic_formula()` method is similar to the
    /// [`Worksheet::write_dynamic_array_formula()`] method, shown above, except
    /// that it writes a dynamic array formula to a single cell, rather than a
    /// range. This is a syntactic shortcut since the array range isn't
    /// generally known for a dynamic range and specifying the initial cell is
    /// sufficient for Excel.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or
    ///   [`Formula`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    pub fn write_dynamic_formula(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: impl Into<Formula>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(row, col, row, col, formula.into(), None, true)
    }

    /// Write a formatted dynamic formula to a worksheet cell.
    ///
    /// The `write_dynamic_formula_with_format()` method is similar to the
    /// [`Worksheet::write_dynamic_array_formula_with_format()`] method, shown
    /// above, except that it writes a dynamic array formula to a single cell,
    /// rather than a range. This is a syntactic shortcut since the array range
    /// isn't generally known for a dynamic range and specifying the initial
    /// cell is sufficient for Excel.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or
    ///   [`Formula`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    pub fn write_dynamic_formula_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: impl Into<Formula>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_array_formula(row, col, row, col, formula.into(), Some(format), true)
    }

    /// Write a blank formatted worksheet cell.
    ///
    /// Write a blank cell with formatting to a worksheet cell. The format is
    /// set via a [`Format`] struct.
    ///
    /// Excel differentiates between an “Empty” cell and a “Blank” cell. An
    /// “Empty” cell is a cell which doesn’t contain data or formatting whilst a
    /// “Blank” cell doesn’t contain data but does contain formatting. Excel
    /// stores “Blank” cells but ignores “Empty” cells.
    ///
    /// The most common case for a formatted blank cell is to write a background
    /// or a border, see the example below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a blank cell with formatting,
    /// i.e., a cell that has no data but does have formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_blank.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatBorder, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     let format1 = Format::new().set_background_color(Color::Yellow);
    ///
    ///     let format2 = Format::new()
    ///         .set_background_color(Color::Yellow)
    ///         .set_border(FormatBorder::Thin);
    ///
    ///     worksheet.write_blank(1, 1, &format1)?;
    ///     worksheet.write_blank(3, 1, &format2)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_blank.png">
    ///
    pub fn write_blank(
        &mut self,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_blank(row, col, format)
    }

    /// Write a url/hyperlink to a worksheet cell.
    ///
    /// Write a url/hyperlink to a worksheet cell with the default Excel
    /// "Hyperlink" cell style.
    ///
    /// There are 3 types of url/link supported by Excel:
    ///
    /// 1. Web based URIs like:
    ///
    ///    * `http://`, `https://`, `ftp://`, `ftps://` and `mailto:`.
    ///
    /// 2. Local file links using the `file://` URI.
    ///
    ///    * `file:///Book2.xlsx`
    ///    * `file:///..\Sales\Book2.xlsx`
    ///    * `file:///C:\Temp\Book1.xlsx`
    ///    * `file:///Book2.xlsx#Sheet1!A1`
    ///    * `file:///Book2.xlsx#'Sales Data'!A1:G5`
    ///
    ///    Most paths will be relative to the root folder, following the Windows
    ///    convention, so most paths should start with `file:///`. For links to
    ///    other Excel files the url string can include a sheet and cell
    ///    reference after the `"#"` anchor, as shown in the last 2 examples
    ///    above. When using Windows paths, like in the examples above, it is
    ///    best to use a Rust raw string to avoid issues with the backslashes:
    ///    `r"file:///C:\Temp\Book1.xlsx"`.
    ///
    /// 3. Internal links to a cell or range of cells in the workbook using the
    ///    pseudo-uri `internal:`:
    ///
    ///    * `internal:Sheet2!A1`
    ///    * `internal:Sheet2!A1:G5`
    ///    * `internal:'Sales Data'!A1`
    ///
    ///    Worksheet references are typically of the form `Sheet1!A1` where a
    ///    worksheet and target cell should be specified. You can also link to a
    ///    worksheet range using the standard Excel range notation like
    ///    `Sheet1!A1:B2`. Excel requires that worksheet names containing spaces
    ///    or non alphanumeric characters are single quoted as follows `'Sales
    ///    Data'!A1`.
    ///
    /// The function will escape the following characters in URLs as required by
    /// Excel, ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
    /// style escapes. In which case it is assumed that the URL was escaped
    /// correctly by the user and will by passed directly to Excel.
    ///
    /// Excel has a limit of around 2080 characters in the url string. Strings
    /// beyond this limit will raise an error, see below.
    ///
    /// For other variants of this function see:
    ///
    /// - [`Worksheet::write_url_with_text()`] to add alternative text to the
    ///   link.
    /// - [`Worksheet::write_url_with_format()`] to add an alternative format to
    ///   the link.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `string`: The url string to write to the cell.
    /// - `link`: The url/hyperlink to write to the cell as a string or [`Url`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxUrlLengthExceeded`] - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - [`XlsxError::UnknownUrlType`] - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - [`XlsxError::ParameterError`] - [`Url`] mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates several of the url writing methods.
    ///
    /// ```
    /// # // This code is available in examples/app_hyperlinks.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError, FormatUnderline};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Create a format to use in the worksheet.
    /// #     let link_format = Format::new()
    /// #         .set_font_color(Color::Red)
    /// #         .set_underline(FormatUnderline::Single);
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet1 = workbook.add_worksheet();
    /// #
    /// #     // Set the column width for clarity.
    /// #     worksheet1.set_column_width(0, 26)?;
    /// #
    ///     // Write some url links.
    ///     worksheet1.write_url(0, 0, "https://www.rust-lang.org")?;
    ///     worksheet1.write_url_with_text(1, 0, "https://www.rust-lang.org", "Learn Rust")?;
    ///     worksheet1.write_url_with_format(2, 0, "https://www.rust-lang.org", &link_format)?;
    ///
    ///     // Write some internal links.
    ///     worksheet1.write_url(4, 0, "internal:Sheet1!A1")?;
    ///     worksheet1.write_url(5, 0, "internal:Sheet2!C4")?;
    ///
    ///     // Write some external links.
    ///     worksheet1.write_url(7, 0, r"file:///C:\Temp\Book1.xlsx")?;
    ///     worksheet1.write_url(8, 0, r"file:///C:\Temp\Book1.xlsx#Sheet1!C4")?;
    ///
    ///     // Add another sheet to link to.
    ///     let worksheet2 = workbook.add_worksheet();
    ///     worksheet2.write_string(3, 2, "Here I am")?;
    ///     worksheet2.write_url_with_text(4, 2, "internal:Sheet1!A6", "Go back")?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("hyperlinks.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/app_hyperlinks.png">
    ///
    /// You can also write the url using a [`Url`] struct:
    ///
    /// ```
    /// # // This code is available in examples/doc_url_intro2.rs
    /// #
    /// # use rust_xlsxwriter::{Url, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write a url with a Url struct.
    ///     worksheet.write_url(0, 0, Url::new("https://www.rust-lang.org"))?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/url_intro1.png">
    ///
    pub fn write_url(
        &mut self,
        row: RowNum,
        col: ColNum,
        link: impl Into<Url>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_url(row, col, &link.into(), None)
    }

    /// Write a url/hyperlink to a worksheet cell with an alternative text.
    ///
    /// Write a url/hyperlink to a worksheet cell with an alternative, user
    /// friendly, text and the default Excel "Hyperlink" cell style.
    ///
    /// This method is similar to [`Worksheet::write_url()`]  except that you
    /// can specify an alternative string for the url. For example you could
    /// have a cell contain the link [Learn Rust](https://www.rust-lang.org)
    /// instead of the raw link <https://www.rust-lang.org>.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `link`: The url/hyperlink to write to the cell as a string or [`Url`].
    /// - `text`: The alternative string to write to the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - Text string exceeds Excel's
    ///   limit of 32,767 characters.
    /// - [`XlsxError::MaxUrlLengthExceeded`] - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - [`XlsxError::UnknownUrlType`] - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - [`XlsxError::ParameterError`] - [`Url`] mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    ///
    /// # Examples
    ///
    /// A simple, getting started, example of some of the features of the
    /// `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_url_with_text.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook , XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write a url and alternative text.
    ///     worksheet.write_url_with_text(0, 0, "https://www.rust-lang.org", "Learn Rust")?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_url_with_text.png">
    ///
    /// You can also write the url using a [`Url`] struct:
    ///
    /// ```
    /// # // This code is available in examples/doc_url_set_text.rs
    /// #
    /// # use rust_xlsxwriter::{Url, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write a url with a Url struct and alternative text.
    ///     worksheet.write(0, 0, Url::new("https://www.rust-lang.org").set_text("Learn Rust"))?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn write_url_with_text(
        &mut self,
        row: RowNum,
        col: ColNum,
        link: impl Into<Url>,
        text: impl Into<String>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        let link = link.into().set_text(text.into());
        self.store_url(row, col, &link, None)
    }

    /// Write a url/hyperlink to a worksheet cell with a user defined format
    ///
    /// Write a url/hyperlink to a worksheet cell with a user defined format
    /// instead of the default Excel "Hyperlink" cell style.
    ///
    /// This method is similar to [`Worksheet::write_url()`] except that you can
    /// specify an alternative format for the url.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `link`: The url/hyperlink to write to the cell as a string or [`Url`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxUrlLengthExceeded`] - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - [`XlsxError::UnknownUrlType`] - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - [`XlsxError::ParameterError`] - [`Url`] mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a url with alternative
    /// format.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_url_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError, FormatUnderline};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a format to use in the worksheet.
    ///     let link_format = Format::new()
    ///         .set_font_color(Color::Red)
    ///         .set_underline(FormatUnderline::Single);
    ///
    ///     // Write a url with an alternative format.
    ///     worksheet.write_url_with_format(0, 0, "https://www.rust-lang.org", &link_format)?;
    /// #
    /// #    // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_url_with_format.png">
    ///
    pub fn write_url_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        link: impl Into<Url>,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_url(row, col, &link.into(), Some(format))
    }

    #[doc(hidden)] // Hide the docs since this is more easily done with a Url struct.
    ///
    /// Write a url/hyperlink to a worksheet cell with various options
    ///
    /// This method is similar to [`Worksheet::write_url()`] and variant methods
    /// except that you can also add a screen tip message, if required.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `link`: The url/hyperlink to write to the cell as a string or [`Url`].
    /// - `text`: The alternative string to write to the cell.
    /// - `tip`: The screen tip string to display when the user hovers over the
    ///   url cell.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// The `text` and `tip` parameters are optional and can be set as a blank
    /// string. The `format` is an `Option<>` parameter and can be specified as
    /// `None` if not required.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - Text string exceeds Excel's
    ///   limit of 32,767 characters.
    /// - [`XlsxError::MaxUrlLengthExceeded`] - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters or the screen tip exceed 255
    ///   characters.
    /// - [`XlsxError::UnknownUrlType`] - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - [`XlsxError::ParameterError`] - [`Url`] mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    ///
    pub fn write_url_with_options(
        &mut self,
        row: RowNum,
        col: ColNum,
        link: impl Into<Url>,
        text: impl Into<String>,
        tip: impl Into<String>,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        let link = link.into().set_text(text.into()).set_tip(tip.into());
        self.store_url(row, col, &link, format)
    }

    /// Write a formatted date and/or time to a worksheet cell.
    ///
    /// The method method writes dates/times that implements [`IntoExcelDateTime`]
    /// to a worksheet cell.
    ///
    /// The date/time types supported are:
    /// - [`ExcelDateTime`].
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - [`chrono::NaiveDateTime`].
    /// - [`chrono::NaiveDate`].
    /// - [`chrono::NaiveTime`].
    ///
    /// [`chrono::NaiveDate`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    /// [`chrono::NaiveTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    /// [`chrono::NaiveDateTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// [`Format`] struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `datetime`: A date/time instance that implements [`IntoExcelDateTime`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted datetimes in an
    /// Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_datetime_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
    ///     let format4 = Format::new().set_num_format("ddd dd mmm yyyy hh:mm");
    ///     let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy hh:mm");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 0)?;
    ///
    ///     // Write the datetime with different Excel formats.
    ///     worksheet.write_datetime_with_format(0, 0, &datetime, &format1)?;
    ///     worksheet.write_datetime_with_format(1, 0, &datetime, &format2)?;
    ///     worksheet.write_datetime_with_format(2, 0, &datetime, &format3)?;
    ///     worksheet.write_datetime_with_format(3, 0, &datetime, &format4)?;
    ///     worksheet.write_datetime_with_format(4, 0, &datetime, &format5)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_datetime.png">
    ///
    /// The following example demonstrates writing formatted dates in an Excel
    /// worksheet.
    ///
    /// ```
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the dates below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd");
    ///     let format4 = Format::new().set_num_format("ddd dd mmm yyyy");
    ///     let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a date object.
    ///     let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
    ///
    ///     // Write the date with different Excel formats.
    ///     worksheet.write_datetime_with_format(0, 0, &date, &format1)?;
    ///     worksheet.write_datetime_with_format(1, 0, &date, &format2)?;
    ///     worksheet.write_datetime_with_format(2, 0, &date, &format3)?;
    ///     worksheet.write_datetime_with_format(3, 0, &date, &format4)?;
    ///     worksheet.write_datetime_with_format(4, 0, &date, &format5)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_date.png">
    ///
    /// The following example demonstrates writing formatted times in an Excel
    /// worksheet.
    ///
    /// ```
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the times below.
    ///     let format1 = Format::new().set_num_format("h:mm");
    ///     let format2 = Format::new().set_num_format("hh:mm");
    ///     let format3 = Format::new().set_num_format("hh:mm:ss");
    ///     let format4 = Format::new().set_num_format("hh:mm:ss.000");
    ///     let format5 = Format::new().set_num_format("h:mm AM/PM");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a time object.
    ///     let time = ExcelDateTime::from_hms_milli(2, 59, 3, 456)?;
    ///
    ///     // Write the time with different Excel formats.
    ///     worksheet.write_datetime_with_format(0, 0, &time, &format1)?;
    ///     worksheet.write_datetime_with_format(1, 0, &time, &format2)?;
    ///     worksheet.write_datetime_with_format(2, 0, &time, &format3)?;
    ///     worksheet.write_datetime_with_format(3, 0, &time, &format4)?;
    ///     worksheet.write_datetime_with_format(4, 0, &time, &format5)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_time.png">
    ///
    ///
    pub fn write_datetime_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        datetime: impl IntoExcelDateTime,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let datetime = datetime.to_excel_serial_date();

        // Store the cell data.
        self.store_datetime(row, col, datetime, Some(format))
    }

    /// Write an unformatted date and/or time to a worksheet cell.
    ///
    /// In general an unformatted date/time isn't very useful since a date in
    /// Excel without a format is just a number. However, this method is
    /// provided for cases where an implicit format is derived from the column
    /// or row format.
    ///
    /// However, for most use cases you should use the
    /// [`write_datetime_with_format()`][Worksheet::write_datetime_with_format]
    /// method with an explicit format.
    ///
    /// The date/time types supported are:
    /// - [`ExcelDateTime`].
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - [`chrono::NaiveDateTime`].
    /// - [`chrono::NaiveDate`].
    /// - [`chrono::NaiveTime`].
    ///
    /// [`chrono::NaiveDate`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    /// [`chrono::NaiveTime`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    /// [`chrono::NaiveDateTime`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `datetime`: A date/time instance that implements
    ///   [`IntoExcelDateTime`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing datetimes that take an
    /// implicit format from the column formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_datetime.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some formats to use with the datetimes below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy hh:mm");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy hh:mm");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
    ///
    ///     // Set the column formats.
    ///     worksheet.set_column_format(0, &format1)?;
    ///     worksheet.set_column_format(1, &format2)?;
    ///     worksheet.set_column_format(2, &format3)?;
    ///
    ///     // Set the column widths for clarity.
    ///     worksheet.set_column_width(0, 20)?;
    ///     worksheet.set_column_width(1, 20)?;
    ///     worksheet.set_column_width(2, 20)?;
    ///
    ///     // Create a datetime object.
    ///     let datetime = ExcelDateTime::from_ymd(2023, 1, 25)?.and_hms(12, 30, 0)?;
    ///
    ///     // Write the datetime without a formats. The dates will get the column
    ///     // format instead.
    ///     worksheet.write_datetime(0, 0, &datetime)?;
    ///     worksheet.write_datetime(0, 1, &datetime)?;
    ///     worksheet.write_datetime(0, 2, &datetime)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_write_datetime_v2.png">
    ///
    pub fn write_datetime(
        &mut self,
        row: RowNum,
        col: ColNum,
        datetime: impl IntoExcelDateTime,
    ) -> Result<&mut Worksheet, XlsxError> {
        let datetime = datetime.to_excel_serial_date();

        // Store the cell data.
        self.store_datetime(row, col, datetime, None)
    }

    #[doc(hidden)] // Hide the docs since this functionality is provided by `write_datetime_with_format()`.
    /// Write a formatted date to a worksheet cell.
    ///
    /// The method method writes dates/times that implements [`IntoExcelDateTime`]
    /// to a worksheet cell.
    ///
    /// The date/time types supported are:
    /// - [`ExcelDateTime`].
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - [`chrono::NaiveDateTime`].
    /// - [`chrono::NaiveDate`].
    /// - [`chrono::NaiveTime`].
    ///
    /// [`chrono::NaiveDate`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    /// [`chrono::NaiveTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    /// [`chrono::NaiveDateTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// [`Format`] struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `date`: A date/time instance that implements [`IntoExcelDateTime`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted dates in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_date.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the dates below.
    ///     let format1 = Format::new().set_num_format("dd/mm/yyyy");
    ///     let format2 = Format::new().set_num_format("mm/dd/yyyy");
    ///     let format3 = Format::new().set_num_format("yyyy-mm-dd");
    ///     let format4 = Format::new().set_num_format("ddd dd mmm yyyy");
    ///     let format5 = Format::new().set_num_format("dddd, mmmm dd, yyyy");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a date object.
    ///     let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
    ///
    ///     // Write the date with different Excel formats.
    ///     worksheet.write_date_with_format(0, 0, &date, &format1)?;
    ///     worksheet.write_date_with_format(1, 0, &date, &format2)?;
    ///     worksheet.write_date_with_format(2, 0, &date, &format3)?;
    ///     worksheet.write_date_with_format(3, 0, &date, &format4)?;
    ///     worksheet.write_date_with_format(4, 0, &date, &format5)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_date.png">
    ///
    pub fn write_date_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        date: impl IntoExcelDateTime,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let datetime = date.to_excel_serial_date();

        // Store the cell data.
        self.store_datetime(row, col, datetime, Some(format))
    }

    #[doc(hidden)] // Hide the docs since this functionality is provided by `write_datetime_with_format()`.
    /// Write a formatted time to a worksheet cell.
    ///
    /// The method method writes dates/times that implements [`IntoExcelDateTime`]
    /// to a worksheet cell.
    ///
    /// The date/time types supported are:
    /// - [`ExcelDateTime`].
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - [`chrono::NaiveDateTime`].
    /// - [`chrono::NaiveDate`].
    /// - [`chrono::NaiveTime`].
    ///
    /// [`chrono::NaiveDate`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    /// [`chrono::NaiveTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    /// [`chrono::NaiveDateTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// [`Format`] struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `time`: A date/time instance that implements [`IntoExcelDateTime`].
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted times in an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_time.rs
    /// #
    /// # use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create some formats to use with the times below.
    ///     let format1 = Format::new().set_num_format("h:mm");
    ///     let format2 = Format::new().set_num_format("hh:mm");
    ///     let format3 = Format::new().set_num_format("hh:mm:ss");
    ///     let format4 = Format::new().set_num_format("hh:mm:ss.000");
    ///     let format5 = Format::new().set_num_format("h:mm AM/PM");
    ///
    ///     // Set the column width for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Create a time object.
    ///     let time = ExcelDateTime::from_hms_milli(2, 59, 3, 456)?;
    ///
    ///     // Write the time with different Excel formats.
    ///     worksheet.write_time_with_format(0, 0, &time, &format1)?;
    ///     worksheet.write_time_with_format(1, 0, &time, &format2)?;
    ///     worksheet.write_time_with_format(2, 0, &time, &format3)?;
    ///     worksheet.write_time_with_format(3, 0, &time, &format4)?;
    ///     worksheet.write_time_with_format(4, 0, &time, &format5)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_time.png">
    ///
    pub fn write_time_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        time: impl IntoExcelDateTime,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let datetime = time.to_excel_serial_date();

        // Store the cell data.
        self.store_datetime(row, col, datetime, Some(format))
    }

    /// Write an unformatted boolean value to a cell.
    ///
    /// Write an unformatted Excel boolean value to a worksheet cell.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to write to the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing boolean values to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_boolean.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_boolean(0, 0, true)?;
    ///     worksheet.write_boolean(1, 0, false)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_boolean.png">
    ///
    pub fn write_boolean(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_boolean(row, col, boolean, None)
    }

    /// Write a formatted boolean value to a worksheet cell.
    ///
    /// Write a boolean value with formatting to a worksheet cell. The format is set
    /// via a [`Format`] struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to write to the cell.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing formatted boolean values to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_write_boolean_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write_boolean_with_format(0, 0, true, &bold)?;
    ///     worksheet.write_boolean_with_format(1, 0, false, &bold)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_write_boolean_with_format.png">
    ///
    ///
    pub fn write_boolean_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Store the cell data.
        self.store_boolean(row, col, boolean, Some(format))
    }

    /// Merge a range of cells.
    ///
    /// The `merge_range()` method allows cells to be merged together so that
    /// they act as a single area.
    ///
    /// The `merge_range()` method writes a string to the merged cells. In order
    /// to write other data types, such as a number or a formula, you can
    /// overwrite the first cell with a call to one of the other
    /// `worksheet.write_*()` functions. The same [`Format`] instance should be
    /// used as was used in the merged range, see the example below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `string`: The string to write to the cell. Other types can also be
    ///   handled. See the documentation above and the example below.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    /// - [`XlsxError::MergeRangeSingleCell`] - A merge range cannot be a single
    ///   cell in Excel.
    /// - [`XlsxError::MergeRangeOverlaps`] - The merge range overlaps a
    ///   previous merge range.
    ///
    ///
    /// # Examples
    ///
    /// An example of creating merged ranges in a worksheet using the
    /// `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/app_merge_range.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, FormatAlign, FormatBorder, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write some merged cells with centering.
    ///     let format = Format::new().set_align(FormatAlign::Center);
    ///
    ///     worksheet.merge_range(1, 1, 1, 2, "Merged cells", &format)?;
    ///
    ///     // Write some merged cells with centering and a border.
    ///     let format = Format::new()
    ///         .set_align(FormatAlign::Center)
    ///         .set_border(FormatBorder::Thin);
    ///
    ///     worksheet.merge_range(3, 1, 3, 2, "Merged cells", &format)?;
    ///
    ///     // Write some merged cells with a number by overwriting the first cell in
    ///     // the string merge range with the formatted number.
    ///     worksheet.merge_range(5, 1, 5, 2, "", &format)?;
    ///     worksheet.write_number_with_format(5, 1, 12345.67, &format)?;
    ///
    ///     // Example with a more complex format and larger range.
    ///     let format = Format::new()
    ///         .set_align(FormatAlign::Center)
    ///         .set_align(FormatAlign::VerticalCenter)
    ///         .set_border(FormatBorder::Thin)
    ///         .set_background_color(Color::Silver);
    ///
    ///     worksheet.merge_range(7, 1, 8, 3, "Merged cells", &format)?;
    /// #
    /// #    // Save the file to disk.
    /// #     workbook.save("merge_range.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/app_merge_range.png">
    ///
    pub fn merge_range(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        string: &str,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions(first_row, first_col)
            || !self.check_dimensions(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Check that the range isn't a singe cell, which isn't allowed by Excel.
        if first_row == last_row && first_col == last_col {
            return Err(XlsxError::MergeRangeSingleCell);
        }

        // Write the first cell in the range.
        self.write_string_with_format(first_row, first_col, string, format)?;

        // Pad out the rest of the range with formatted blanks cells. We split
        // this into the first row and subsequent rows to allow us to handle
        // "constant mode" write-ahead.
        for col in first_col + 1..=last_col {
            self.write_blank(first_row, col, format)?;
        }

        self.set_writing_ahead(true);
        for row in first_row + 1..=last_row {
            for col in first_col..=last_col {
                self.write_blank(row, col, format)?;
            }
        }
        self.set_writing_ahead(false);

        // Create a cell range for storage and range testing.
        let cell_range = CellRange::new(first_row, first_col, last_row, last_col);

        // Check if the merged range overlaps any previous merged range. This is
        // a major error in Excel. Note, the ranges are stored in a separate Vec
        // to the cells to cut down on storage size.
        let new_index = self.merged_ranges.len();
        for row in first_row..=last_row {
            for col in first_col..=last_col {
                match self.merged_cells.get_mut(&(row, col)) {
                    Some(index) => {
                        let previous_cell_range = self.merged_ranges.get(*index).unwrap();
                        return Err(XlsxError::MergeRangeOverlaps(
                            cell_range.to_error_string(),
                            previous_cell_range.to_error_string(),
                        ));
                    }
                    None => self.merged_cells.insert((row, col), new_index),
                };
            }
        }

        // Store the merge range if everything was okay.
        self.merged_ranges.push(cell_range);

        Ok(self)
    }

    /// Add an image to a worksheet.
    ///
    /// Add an image to a worksheet at a cell location. The image should be
    /// encapsulated in an [`Image`] object.
    ///
    /// The supported image formats are:
    ///
    /// - PNG
    /// - JPG
    /// - GIF: The image can be an animated gif in more recent versions of
    ///   Excel.
    /// - BMP: BMP images are only supported for backward compatibility. In
    ///   general it is best to avoid BMP images since they are not compressed.
    ///   If used, BMP images must be 24 bit, true color, bitmaps.
    ///
    /// EMF and WMF file formats will be supported in an upcoming version of the
    /// library.
    ///
    /// **NOTE on SVG files**: Excel doesn't directly support SVG files in the
    /// same way as other image file formats. It allows SVG to be inserted into
    /// a worksheet but converts them to, and displays them as, PNG files. It
    /// stores the original SVG image in the file so the original format can be
    /// retrieved. This removes the file size and resolution advantage of using
    /// SVG files. As such SVG files are not supported by `rust_xlsxwriter`
    /// since a conversion to the PNG format would be required and that format
    /// is already supported.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The [`Image`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a new Image object and
    /// adding it to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_image.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new image object.
    ///     let image = Image::new("examples/rust_logo.png")?;
    ///
    ///     // Insert the image.
    ///     worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/image_intro.png">
    ///
    pub fn insert_image(
        &mut self,
        row: RowNum,
        col: ColNum,
        image: &Image,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.insert_image_with_offset(row, col, image, 0, 0)?;

        Ok(self)
    }

    /// Add an image to a worksheet at an offset.
    ///
    /// Add an image to a worksheet at a pixel offset within a cell location.
    /// The image should be encapsulated in an [`Image`] object.
    ///
    /// This method is similar to [`Worksheet::insert_image()`] except that the
    /// image can be offset from the top left of the cell.
    ///
    /// Note, it is possible to offset the image outside the target cell if
    /// required.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The [`Image`] to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// This example shows how to add an image to a worksheet at an offset
    /// within the cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_image_with_offset.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new image object.
    ///     let image = Image::new("examples/rust_logo.png")?;
    ///
    ///     // Insert the image at an offset.
    ///     worksheet.insert_image_with_offset(1, 2, &image, 10, 5)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_insert_image_with_offset.png">
    ///
    pub fn insert_image_with_offset(
        &mut self,
        row: RowNum,
        col: ColNum,
        image: &Image,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        let mut image = image.clone();
        image.x_offset = x_offset;
        image.y_offset = y_offset;

        // Store workbook objects in row by column position order.
        self.images.insert((row, col, y_offset, x_offset), image);

        Ok(self)
    }

    /// Embed an image to a worksheet and fit it to a cell.
    ///
    /// This method can be used to embed a image into a worksheet cell and have
    /// the image automatically scale to the width and height of the cell. The
    /// X/Y scaling of the image is preserved but the size of the image is
    /// adjusted to fit the largest possible width or height depending on the
    /// cell dimensions.
    ///
    /// This is the equivalent of Excel's menu option to insert an image using
    /// the option to "Place in Cell" which is only available in Excel 365
    /// versions from 2023 onwards. For older versions of Excel a `#VALUE!`
    /// error is displayed.
    ///
    /// The image should be encapsulated in an [`Image`] object. See
    /// [`Worksheet::insert_image()`] above for details on the supported image
    /// types.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The [`Image`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::ParameterError`] - Embedded images can only be added to
    ///   the current row in "constant memory" mode. They cannot be added to a
    ///   previously written row.
    ///
    /// # Examples
    ///
    /// An example of embedding images into a worksheet cells using
    /// `rust_xlsxwriter`. This image scales to size of the cell and moves with
    /// it.
    ///
    /// This is the equivalent of Excel's menu option to insert an image using
    /// the option to "Place in Cell".
    ///
    /// ```
    /// # // This code is available in examples/app_embedded_images.rs
    /// #
    /// use rust_xlsxwriter::{Image, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create a new image object.
    ///     let image = Image::new("examples/rust_logo.png")?;
    ///
    ///     // Widen the first column to make the caption clearer.
    ///     worksheet.set_column_width(0, 30)?;
    ///     worksheet.write(0, 0, "Embed images that scale to the cell size")?;
    ///
    ///     // Change cell widths/heights to demonstrate the image differences.
    ///     worksheet.set_column_width(1, 14)?;
    ///     worksheet.set_row_height(1, 60)?;
    ///     worksheet.set_row_height(3, 90)?;
    ///
    ///     // Embed the images in cells of different widths/heights.
    ///     worksheet.embed_image(1, 1, &image)?;
    ///     worksheet.embed_image(3, 1, &image)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("embedded_images.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/embedded_images.png">
    ///
    pub fn embed_image(
        &mut self,
        row: RowNum,
        col: ColNum,
        image: &Image,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.store_embedded_image(row, col, image, None)
    }

    /// Embed an image to a worksheet and fit it to a formatted cell.
    ///
    /// This method can be used to embed a image into a worksheet cell and have
    /// the image automatically scale to the width and height of the cell. This
    /// is similar to the [`Worksheet::embed_image()`] above but it allows you
    /// to add an additional cell format using [`Format`]. This is occasionally
    /// useful if you want to set a cell border around the image or a cell
    /// background color.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The [`Image`] to insert into the cell.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    pub fn embed_image_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        image: &Image,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.store_embedded_image(row, col, image, Some(format))
    }

    /// Add an image to a worksheet and fit it to a cell.
    ///
    /// Add an image to a worksheet and scale it so that it fits in a cell. This
    /// is similar in effect to [`Worksheet::embed_image()`] but in Excel's
    /// terminology it inserts the image placed *over* the cell instead of *in*
    /// the cell. The only advantage of this method is that the output file will
    /// work will all versions of Excel. The `Worksheet::embed_image()` method
    /// only works with versions of Excel from 2003 onwards.
    ///
    /// This method can be useful when creating a product spreadsheet with a
    /// column of images for each product. The image should be encapsulated in
    /// an [`Image`] object. See [`Worksheet::insert_image()`] above for details
    /// on the supported image types. The scaling calculation for this method
    /// takes into account the DPI of the image in the same way that Excel does.
    ///
    /// There are two options, which are controlled by the `keep_aspect_ratio`
    /// parameter. The image can be scaled vertically and horizontally to occupy
    /// the entire cell or the aspect ratio of the image can be maintained so
    /// that the image is scaled to the lesser of the horizontal or vertical
    /// sizes. See the example below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The [`Image`] to insert into the cell.
    /// - `keep_aspect_ratio`: Boolean value to maintain the aspect ratio of the
    ///   image if `true` or scale independently in the horizontal and vertical
    ///   directions if `false`.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// An example of inserting images into a worksheet using `rust_xlsxwriter`
    /// so that they are scaled to a cell. This approach can be useful if you
    /// are building up a spreadsheet of products with a column of images for
    /// each product.
    ///
    /// ```
    /// # // This code is available in examples/app_images_fit_to_cell.rs
    /// #
    /// # use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let center = Format::new().set_align(FormatAlign::VerticalCenter);
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Widen the first column to make the text clearer.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Set larger cells to accommodate the images.
    ///     worksheet.set_column_width_pixels(1, 200)?;
    ///     worksheet.set_row_height_pixels(0, 140)?;
    ///     worksheet.set_row_height_pixels(2, 140)?;
    ///     worksheet.set_row_height_pixels(4, 140)?;
    ///
    ///     // Create a new image object.
    ///     let image = Image::new("examples/rust_logo.png")?;
    ///
    ///     // Insert the image as standard, without scaling.
    ///     worksheet.write_with_format(0, 0, "Unscaled image inserted into cell:", &center)?;
    ///     worksheet.insert_image(0, 1, &image)?;
    ///
    ///     // Insert the image and scale it to fit the entire cell.
    ///     worksheet.write_with_format(2, 0, "Image scaled to fit cell:", &center)?;
    ///     worksheet.insert_image_fit_to_cell(2, 1, &image, false)?;
    ///
    ///     // Insert the image and scale it to the cell while maintaining the aspect ratio.
    ///     // In this case it is scaled to the smaller of the width or height scales.
    ///     worksheet.write_with_format(4, 0, "Image scaled with a fixed aspect ratio:", &center)?;
    ///     worksheet.insert_image_fit_to_cell(4, 1, &image, true)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("images_fit_to_cell.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_images_fit_to_cell.png">
    ///
    pub fn insert_image_fit_to_cell(
        &mut self,
        row: RowNum,
        col: ColNum,
        image: &Image,
        keep_aspect_ratio: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        let width = self.column_pixel_width(col, image.object_movement);
        let height = self.row_pixel_height(row, image.object_movement);

        let image = image
            .clone()
            .set_scale_to_size(width, height, keep_aspect_ratio);

        self.images.insert((row, col, 0, 0), image);

        Ok(self)
    }

    /// Insert a background image into a worksheet.
    ///
    /// A background image can be added to a worksheet to add a watermark or
    /// display a company logo. Excel repeats the image for the entirety of the
    /// worksheet.
    ///
    /// The image should be encapsulated in an [`Image`] object. See
    /// [`Worksheet::insert_image()`] above for details on the supported image
    /// types.
    ///
    /// As an alternative to background images, it should be noted that the
    /// Microsoft Excel documentation recommends setting a watermark via an
    /// image in the worksheet header. An example of that technique is shown in
    /// the [`Worksheet::set_header_image()`] examples.
    ///
    /// # Parameters
    ///
    /// - `image`: The [`Image`] to use as the worksheet background.
    ///
    /// # Examples
    ///
    /// An example of inserting a background image into a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/app_background_image.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // The image may not be visible unless the view is large.
    /// #     worksheet.write(0, 0, "Scroll down and right to see the background image")?;
    /// #
    ///     // Create a new image object.
    ///     let image = Image::new("examples/watermark.png")?;
    ///
    ///     // Insert the background image.
    ///     worksheet.insert_background_image(&image);
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("background_image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_background_image.png">
    ///
    pub fn insert_background_image(&mut self, image: &Image) -> &mut Worksheet {
        self.background_image = Some(image.clone());
        self
    }

    /// Add a chart to a worksheet.
    ///
    /// Add a [`Chart`] to a worksheet at a cell location.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `chart`: The [`Chart`] to insert into the cell.
    ///
    /// When used with a [Chartsheet](Worksheet::new_chartsheet) the row/column
    /// arguments are ignored but it is best to use `(0, 0)` for clarity.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::ChartError`] - A general error that is raised when a
    ///   chart parameter is incorrect or a chart is configured incorrectly.
    ///
    /// # Examples
    ///
    /// Insert a chart object into a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_chart_simple.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_simple.png">
    ///
    pub fn insert_chart(
        &mut self,
        row: RowNum,
        col: ColNum,
        chart: &Chart,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.insert_chart_with_offset(row, col, chart, 0, 0)?;

        Ok(self)
    }

    /// Add a chart to a worksheet at an offset.
    ///
    /// Add a [`Chart`] to a worksheet  at a pixel offset within a cell
    /// location.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `chart`: The [`Chart`] to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::ChartError`] - A general error that is raised when a
    ///   chart parameter is incorrect or a chart is configured incorrectly.
    ///
    /// # Examples
    ///
    /// Example of adding a chart to a worksheet with a pixel offset within the
    /// cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_chart_with_offset.rs
    /// #
    /// # use rust_xlsxwriter::{Chart, ChartType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the chart.
    /// #     worksheet.write(0, 0, 50)?;
    /// #     worksheet.write(1, 0, 30)?;
    /// #     worksheet.write(2, 0, 40)?;
    /// #
    /// #     // Create a new chart.
    ///     let mut chart = Chart::new(ChartType::Column);
    ///
    ///     // Add a data series using Excel formula syntax to describe the range.
    ///     chart.add_series().set_values("Sheet1!$A$1:$A$3");
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart_with_offset(0, 2, &chart, 10, 5)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("chart.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_insert_chart_with_offset.png">
    ///
    pub fn insert_chart_with_offset(
        &mut self,
        row: RowNum,
        col: ColNum,
        chart: &Chart,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        let mut chart = chart.clone();
        chart.is_chartsheet = self.is_chartsheet;

        // Check that the chart has been set up correctly.
        chart.validate()?;

        chart.x_offset = x_offset;
        chart.y_offset = y_offset;

        // Limit chartsheets to one chart in position (0, 0).
        let mut row = row;
        let mut col = col;
        if self.is_chartsheet {
            row = 0;
            col = 0;
            chart.x_offset = 0;
            chart.y_offset = 0;
        }

        // Store workbook objects in row by column position order.
        self.charts.insert((row, col, y_offset, x_offset), chart);

        Ok(self)
    }

    /// Add a Note to a cell.
    ///
    /// A Note is a post-it style message that is revealed when the user mouses
    /// over a worksheet cell. The presence of a Note is indicated by a small
    /// red triangle in the upper right-hand corner of the cell.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/doc_note_intro.png">
    ///
    /// In versions of Excel prior to Office 365 Notes were referred to as
    /// "Comments". The name Comment is now used for a newer style threaded
    /// comment and Note is used for the older non threaded version. See the
    /// Microsoft docs on [The difference between threaded comments and notes].
    ///
    /// [The difference between threaded comments and notes]:
    ///     https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3
    ///
    /// See [`Note`] for details on the properties of Notes.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `note`: The [`Note`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - Text exceeds Excel's limit of
    ///   32,713 characters.
    ///
    /// # Examples
    ///
    /// An example of writing cell Notes to a worksheet using the
    /// `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/app_notes.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Widen the first column for clarity.
    /// #     worksheet.set_column_width(0, 16)?;
    /// #
    /// #     // Write some data.
    /// #     let party_items = [
    /// #         "Invitations",
    /// #         "Doors",
    /// #         "Flowers",
    /// #         "Champagne",
    /// #         "Menu",
    /// #         "Peter",
    /// #     ];
    /// #     worksheet.write_column(0, 0, party_items)?;
    /// #
    ///     // Create a new worksheet Note.
    ///     let note = Note::new("I will get the flowers myself").set_author("Clarissa Dalloway");
    ///
    ///     // Add the note to a cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/app_notes.png">
    ///
    pub fn insert_note(
        &mut self,
        row: RowNum,
        col: ColNum,
        note: &Note,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        //  Check that the string is < Excel limit of 32767 chars, - 54
        //  characters to allow for the author name prefix.
        if note.text.chars().count() > MAX_STRING_LEN - 54 {
            return Err(XlsxError::MaxStringLengthExceeded);
        }

        // Set the cell that the Note refers to. This is different form the cell
        // where the note appears.
        let mut note = note.clone();
        note.cell_row = row;
        note.cell_col = col;

        // Store the note in a structure similar to the worksheet data table
        // since notes also affect the calculation of <row> span attributes.
        match self.notes.entry(row) {
            Entry::Occupied(mut entry) => {
                // The row already exists. Insert/replace column value.
                let columns = entry.get_mut();
                columns.insert(col, note);
            }
            Entry::Vacant(entry) => {
                // The row doesn't exist, create a new row with columns and
                // insert the cell value.
                let columns = BTreeMap::from([(col, note)]);
                entry.insert(columns);
            }
        }

        self.has_vml = true;

        Ok(self)
    }

    /// Insert a textbox shape into a worksheet.
    ///
    /// This method can be used to insert an Excel Textbox shape with text into
    /// a worksheet.
    ///
    /// See the [`Shape`] documentation for a detailed description of the
    /// methods that can be used to configure the size and appearance of the
    /// textbox.
    ///
    /// Note, no Excel shape other than Textbox is supported. See [Support for
    /// other Excel shape
    /// types](crate::Shape#support-for-other-excel-shape-types).
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `shape`: The [`Shape`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_shape.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape and add some text.
    ///     let textbox = Shape::textbox().set_text("This is some text");
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_insert_shape.png">
    ///
    pub fn insert_shape(
        &mut self,
        row: RowNum,
        col: ColNum,
        shape: &Shape,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.insert_shape_with_offset(row, col, shape, 0, 0)?;

        Ok(self)
    }

    /// Insert a textbox shape into a worksheet cell at an offset.
    ///
    /// This method can be used to insert an Excel Textbox shape with text into
    /// a worksheet cell at a pixel offset.
    ///
    /// See the [`Shape`] documentation for a detailed description of the
    /// methods that can be used to configure the size and appearance of the
    /// textbox.
    ///
    /// Note, no Excel shape other than Textbox is supported. See [Support for
    /// other Excel shape
    /// types](crate::Shape#support-for-other-excel-shape-types).
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `shape`: The [`Shape`] to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape to a worksheet cell at
    /// an offset.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_shape_with_offset.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape and add some text.
    ///     let textbox = Shape::textbox().set_text("This is some text");
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape_with_offset(1, 1, &textbox, 10, 5)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_insert_shape_with_offset.png">
    ///
    pub fn insert_shape_with_offset(
        &mut self,
        row: RowNum,
        col: ColNum,
        shape: &Shape,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        let mut shape = shape.clone();
        shape.x_offset = x_offset;
        shape.y_offset = y_offset;

        // Store workbook objects in row by column position order.
        self.shapes.insert((row, col, y_offset, x_offset), shape);

        Ok(self)
    }

    /// Make all worksheet notes visible when the file loads.
    ///
    /// By default Excel hides cell notes until the user mouses over the parent
    /// cell. However, if required you can make all worksheet notes visible when
    /// the worksheet loads. You can also make individual notes visible using
    /// the [`Note::set_visible()`](crate::Note::set_visible) method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding notes to a worksheet and
    /// setting the worksheet property to make them all visible.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_show_all_notes.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note");
    ///
    ///     // Add the note to some worksheet cells.
    ///     worksheet.insert_note(1, 0, &note)?;
    ///     worksheet.insert_note(3, 3, &note)?;
    ///     worksheet.insert_note(6, 0, &note)?;
    ///     worksheet.insert_note(8, 3, &note)?;
    ///
    ///     // Display all the notes in the worksheet.
    ///     worksheet.show_all_notes(true);
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_show_all_notes.png">
    ///
    pub fn show_all_notes(&mut self, enable: bool) -> &mut Worksheet {
        self.show_all_notes = enable;
        self
    }

    /// Set the default author name for all the notes in the worksheet.
    ///
    /// The Note author is the creator of the note. In Excel the author name is
    /// taken from the user name in the options/preference dialog. The note
    /// author name appears in two places: at the start of the note text in bold
    /// and at the bottom of the worksheet in the status bar.
    ///
    /// If no name is specified the default name "Author" will be applied to the
    /// note. The author name for individual notes can be set via the
    /// [`Note::set_author()`](crate::Note::set_author) method. Alternatively
    /// this method can be used to set the default author name for all notes in
    /// a worksheet.
    ///
    /// # Parameters
    ///
    /// - `name`: The note author name. Must be less than or equal to the Excel
    ///   limit of 52 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding notes to a worksheet and setting
    /// the default author name.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_default_note_author.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note");
    ///
    ///     // Add the note to some worksheet cells.
    ///     worksheet.insert_note(2, 0, &note)?;
    ///
    ///     // Display all the notes in the worksheet.
    ///     worksheet.set_default_note_author("Rust");
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_default_note_author.png">
    ///
    pub fn set_default_note_author(&mut self, name: impl Into<String>) -> &mut Worksheet {
        let name = name.into();
        if name.chars().count() > 52 {
            eprintln!("Author string must be less than the Excel limit of 52 characters: {name}");
            return self;
        }

        self.note_authors = BTreeMap::from([(name, 0)]);

        self
    }

    /// Add a Excel Form Control button object to a worksheet.
    ///
    /// Add a [`Button`] to a worksheet at a cell location. The worksheet button
    /// object is mainly provided as a way of triggering a VBA macro, see
    /// [Working with VBA macros](crate::macros) for more details.
    ///
    /// Note, Button is the only VBA Control supported by `rust_xlsxwriter`. It
    /// is unlikely that any other Excel form elements will be added in the
    /// future due to the implementation effort required.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `button`: The [`Button`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// An example of adding macros to an `rust_xlsxwriter` file using a VBA
    /// macros file extracted from an existing Excel xlsm file.
    ///
    /// ```
    /// # // This code is available in examples/app_macros.rs
    /// #
    /// use rust_xlsxwriter::{Button, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add the VBA macro file.
    ///     workbook.add_vba_project("examples/vbaProject.bin")?;
    ///
    ///     // Add a worksheet and some text.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Widen the first column for clarity.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     worksheet.write(2, 0, "Press the button to say hello:")?;
    ///
    ///     // Add a button tied to a macro in the VBA project.
    ///     let button = Button::new()
    ///         .set_caption("Press Me")
    ///         .set_macro("say_hello")
    ///         .set_width(80)
    ///         .set_height(30);
    ///
    ///     worksheet.insert_button(2, 1, &button)?;
    ///
    ///     // Save the file to disk. Note the `.xlsm` extension. This is required by
    ///     // Excel or it will raise a warning.
    ///     workbook.save("macros.xlsm")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/app_macros.png">
    ///
    pub fn insert_button(
        &mut self,
        row: RowNum,
        col: ColNum,
        button: &Button,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.insert_button_with_offset(row, col, button, 0, 0)?;

        Ok(self)
    }

    /// Add a Excel Form Control button object to a  at an offset.
    ///
    /// Add a [`Button`] to a worksheet  at a pixel offset within a cell
    /// location. See [`Worksheet::insert_button()`] above
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `button`: The [`Button`] to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    pub fn insert_button_with_offset(
        &mut self,
        row: RowNum,
        col: ColNum,
        button: &Button,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        let mut button = button.clone();
        button.x_offset = x_offset;
        button.y_offset = y_offset;

        // Store workbook objects in row by column position order.
        self.buttons.insert((row, col, y_offset, x_offset), button);
        self.has_vml = true;

        Ok(self)
    }

    /// Insert a boolean checkbox in a worksheet cell.
    ///
    /// Checkboxes are a [new feature] added to Excel in 2024. They are a way of
    /// displaying a boolean value as a checkbox in a cell. The underlying value
    /// is still an Excel `TRUE/FALSE` boolean value and can be used in formulas
    /// and in references.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/checkbox.png">
    ///
    /// [new feature]:
    ///     https://techcommunity.microsoft.com/blog/excelblog/introducing-checkboxes-in-excel/4173561
    ///
    /// The `insert_checkbox()` method can be used to replicate this behavior,
    /// see the examples below.
    ///
    /// The checkbox feature is only available in Excel versions from 2024 and
    /// later. In older versions the value will be displayed as a standard Excel
    /// `TRUE` or `FALSE` boolean. In fact Excel actually stores a checkbox as a
    /// normal boolean but with a special format. If required you can make use
    /// of this property to create a checkbox with
    /// [`Worksheet::write_boolean_with_format()`] and a cell format that has
    /// the [`Format::set_checkbox()`] property set, see the second example
    /// below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to display as a checkbox.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding adding checkbox boolean values to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_checkbox1.rs
    /// #
    /// use rust_xlsxwriter::{Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Insert some boolean checkboxes to the worksheet.
    ///     worksheet.insert_checkbox(2, 2, false)?;
    ///     worksheet.insert_checkbox(3, 2, true)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("worksheet.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_insert_checkbox1.png">
    ///
    /// This example demonstrates adding adding checkboxes by making use of the
    /// Excel feature that a checkbox is actually a boolean value with a special
    /// format.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_checkbox2.rs
    /// #
    /// use rust_xlsxwriter::{Format, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create a checkbox format.
    ///     let format = Format::new().set_checkbox();
    ///
    ///     // Insert some boolean checkboxes to the worksheet.
    ///     worksheet.write_boolean_with_format(2, 2, false, &format)?;
    ///     worksheet.write_boolean_with_format(3, 2, true, &format)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("worksheet.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    /// The produces the same output as the previous example.
    ///
    pub fn insert_checkbox(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Create a checkbox format.
        let format = Format::new().set_checkbox();

        // Store the cell data.
        self.store_boolean(row, col, boolean, Some(&format))
    }

    /// Insert a boolean checkbox in a worksheet cell with a cell format.
    ///
    /// This method allow you to insert a boolean checkbox in a worksheet cell
    /// with a background color or other cell format property.
    ///
    /// See the [`Worksheet::insert_checkbox()`] method above for more details.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to display as a checkbox.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    ///
    /// # Examples
    ///
    /// This example demonstrates adding adding a checkbox boolean value to a
    /// worksheet along with a cell format.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_insert_checkbox_with_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some cell formats with different colors.
    ///     let format1 = Format::new().set_background_color("FFC7CE");
    ///     let format2 = Format::new().set_background_color("C6EFCE");
    ///
    ///     // Insert some boolean checkboxes to the worksheet.
    ///     worksheet.insert_checkbox_with_format(2, 2, false, &format1)?;
    ///     worksheet.insert_checkbox_with_format(3, 2, true, &format2)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_insert_checkbox_with_format.png">
    ///
    pub fn insert_checkbox_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Add the checkbox format to the user format.
        let format = format.clone().set_checkbox();

        // Store the cell data.
        self.store_boolean(row, col, boolean, Some(&format))
    }

    /// Set the height for a row of cells.
    ///
    /// The `set_row_height()` method is used to change the default height of a
    /// row. The height is specified in character units, where the default
    /// height is 15. Excel allows height values in increments of 0.25.
    ///
    /// To specify the height in pixels use the
    /// [`Worksheet::set_row_height_pixels()`] method.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `height`: The row height in character units.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the height for a row in
    /// Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_row_height.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string(0, 0, "Normal")?;
    ///     worksheet.write_string(2, 0, "Taller")?;
    ///
    ///     // Set the row height in Excel character units.
    ///     worksheet.set_row_height(2, 30)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_row_height.png">
    ///
    pub fn set_row_height(
        &mut self,
        row: RowNum,
        height: impl Into<f64>,
    ) -> Result<&mut Worksheet, XlsxError> {
        let height = height.into();

        // If the height is 0 then the Excel treats the row as hidden with
        // default height.
        if height == 0.0 {
            return self.set_row_hidden(row);
        }

        // Set a suitable column range for the row dimension check/set.
        let min_col = self.get_min_col();

        // Check row is in the allowed range.
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing row metadata object or create a new one.
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => row_options.height = height,
            None => {
                let row_options = RowOptions {
                    height,
                    xf_index: 0,
                    hidden: false,
                    level: 0,
                    collapsed: false,
                    format: None,
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    /// Set the height for a row of cells, in pixels.
    ///
    /// The `set_row_height_pixels()` method is used to change the default height of a
    /// row. The height is specified in pixels, where the default
    /// height is 20.
    ///
    /// To specify the height in Excel's character units use the
    /// [`Worksheet::set_row_height()`] method.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `height`: The row height in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the height for a row in Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_row_height_pixels.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string(0, 0, "Normal")?;
    ///     worksheet.write_string(2, 0, "Taller")?;
    ///
    ///     // Set the row height in pixels.
    ///     worksheet.set_row_height_pixels(2, 40)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_row_height.png">
    ///
    pub fn set_row_height_pixels(
        &mut self,
        row: RowNum,
        height: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        let height = 0.75 * f64::from(height);

        self.set_row_height(row, height)
    }

    /// Set the format for a row of cells.
    ///
    /// The `set_row_format()` method is used to change the default format of a
    /// row. Any unformatted data written to that row will then adopt that
    /// format. Formatted data written to the row will maintain its own cell
    /// format. See the example below.
    ///
    /// A future version of this library may support automatic merging of
    /// explicit cell formatting with the row formatting but that isn't
    /// currently supported.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format for a row in Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_row_format.rs
    /// #
    /// # use rust_xlsxwriter::{Color, Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add for formats.
    ///     let bold_format = Format::new().set_bold();
    ///     let red_format = Format::new().set_font_color(Color::Red);
    ///
    ///     // Set the row format.
    ///     worksheet.set_row_format(1, &red_format)?;
    ///
    ///     // Add some unformatted text that adopts the row format.
    ///     worksheet.write_string(1, 0, "Hello")?;
    ///
    ///     // Add some formatted text that overrides the row format.
    ///     worksheet.write_string_with_format(1, 2, "Hello", &bold_format)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_row_format.png">
    ///
    pub fn set_row_format(
        &mut self,
        row: RowNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Get a suitable minimum value for the dimension check/set.
        let min_col = self.get_min_col();

        // Check row is in the allowed range.
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_xf_index(format);

        // Update an existing row metadata object or create a new one.
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => row_options.xf_index = xf_index,
            None => {
                let row_options = RowOptions {
                    height: self.user_default_row_height,
                    xf_index,
                    hidden: false,
                    level: 0,
                    collapsed: false,
                    format: Some(format.clone()),
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    /// Group a range of rows into a worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// A grouping is created as follows:
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_rows_intro1.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add grouping over the sub-total range.
    ///     worksheet.group_rows(1, 10)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Which creates a grouping at level 1:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro1.png">
    ///
    /// Hierarchical sub-groups are created by repeating the method calls for a
    /// sub-range of an upper level group:
    ///
    /// ```rust
    /// # // This code is available in examples/doc_worksheet_group_rows_intro2.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add Level 1 grouping over the sub-total range.
    ///     worksheet.group_rows(1, 10)?;
    ///
    ///     // Add secondary groups within the first range.
    ///     worksheet.group_rows(1, 4)?;
    ///     worksheet.group_rows(6, 9)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// This creates the following grouping and sub-grouping at levels 1 and 2:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro2.png">
    ///
    /// It should be noted that Excel requires outline groups at the same level
    /// to be separated by at least one row (or column) or else it will merge
    /// them into a single group. This is generally to allow a subtotal
    /// row/column.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. Zero indexed.
    /// - `last_row`: The last row of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row. Note, to reverse the group direction see the
    ///   [`Worksheet::group_symbols_above()`] method.
    /// - [`XlsxError::MaxGroupLevelExceeded`] - Group depth level exceeds
    ///   Excel's limit of 8 levels.
    ///
    /// # Examples
    ///
    /// An example of how to group worksheet rows into outlines.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_rows1.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.write_with_format(0, 0, "Region", &bold)?;
    ///     worksheet.write(1, 0, "North 1")?;
    ///     worksheet.write(2, 0, "North 2")?;
    ///     worksheet.write(3, 0, "North 3")?;
    ///     worksheet.write(4, 0, "North 4")?;
    ///     worksheet.write_with_format(5, 0, "North Total", &bold)?;
    ///
    ///     worksheet.write_with_format(0, 1, "Sales", &bold)?;
    ///     worksheet.write(1, 1, 1000)?;
    ///     worksheet.write(2, 1, 1200)?;
    ///     worksheet.write(3, 1, 900)?;
    ///     worksheet.write(4, 1, 1200)?;
    ///     worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", &bold)?;
    ///
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add grouping over the sub-total range.
    ///     worksheet.group_rows(1, 4)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows1.png">
    ///
    /// An example of how to group worksheet rows into outlines. This example
    /// shows hows to add secondary groups within a primary grouping. Excel
    /// requires at least one row between each outline grouping at the same
    /// level.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_rows2.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_with_format(0, 0, "Region", &bold)?;
    /// #     worksheet.write(1, 0, "North 1")?;
    /// #     worksheet.write(2, 0, "North 2")?;
    /// #     worksheet.write(3, 0, "North 3")?;
    /// #     worksheet.write(4, 0, "North 4")?;
    /// #     worksheet.write_with_format(5, 0, "North Total", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 1, "Sales", &bold)?;
    /// #     worksheet.write(1, 1, 1000)?;
    /// #     worksheet.write(2, 1, 1200)?;
    /// #     worksheet.write(3, 1, 900)?;
    /// #     worksheet.write(4, 1, 1200)?;
    /// #     worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", &bold)?;
    /// #
    /// #     worksheet.write(6, 0, "South 1")?;
    /// #     worksheet.write(7, 0, "South 2")?;
    /// #     worksheet.write(8, 0, "South 3")?;
    /// #     worksheet.write(9, 0, "South 4")?;
    /// #     worksheet.write_with_format(10, 0, "South Total", &bold)?;
    /// #
    /// #     worksheet.write(6, 1, 400)?;
    /// #     worksheet.write(7, 1, 600)?;
    /// #     worksheet.write(8, 1, 500)?;
    /// #     worksheet.write(9, 1, 600)?;
    /// #     worksheet.write_formula_with_format(10, 1, "=SUBTOTAL(9,B7:B10)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(11, 0, "Grand Total", &bold)?;
    /// #     worksheet.write_formula_with_format(11, 1, "=SUBTOTAL(9,B2:B11)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add grouping over the sub-total range.
    ///     worksheet.group_rows(1, 10)?;
    ///
    ///     // Add secondary groups within the first range.
    ///     worksheet.group_rows(1, 4)?;
    ///     worksheet.group_rows(6, 9)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows2.png">
    ///
    ///
    ///
    pub fn group_rows(
        &mut self,
        first_row: RowNum,
        last_row: RowNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.set_grouped_rows(first_row, last_row, false)
    }

    /// Group a range of rows into a collapsed worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// See [`Worksheet::group_rows()`] above for an explanation on how to create
    /// sub-groupings.
    ///
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. Zero indexed.
    /// - `last_row`: The last row of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row. Note, to reverse the group direction see the
    ///   [`Worksheet::group_symbols_above()`] method.
    /// - [`XlsxError::MaxGroupLevelExceeded`] - Group depth level exceeds
    ///   Excel's limit of 8 levels.
    ///
    /// # Examples
    ///
    /// An example of how to group worksheet rows into outlines with
    /// collapsed/hidden rows.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_rows_collapsed1.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.write_with_format(0, 0, "Region", &bold)?;
    ///     worksheet.write(1, 0, "North 1")?;
    ///     worksheet.write(2, 0, "North 2")?;
    ///     worksheet.write(3, 0, "North 3")?;
    ///     worksheet.write(4, 0, "North 4")?;
    ///     worksheet.write_with_format(5, 0, "North Total", &bold)?;
    ///
    ///     worksheet.write_with_format(0, 1, "Sales", &bold)?;
    ///     worksheet.write(1, 1, 1000)?;
    ///     worksheet.write(2, 1, 1200)?;
    ///     worksheet.write(3, 1, 900)?;
    ///     worksheet.write(4, 1, 1200)?;
    ///     worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", &bold)?;
    ///
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add collapse grouping over the sub-total range.
    ///     worksheet.group_rows_collapsed(1, 4)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_collapsed1.png">
    ///
    /// For comparison here is the expanded output:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows1.png">
    ///
    /// An example of how to group worksheet rows into outlines with
    /// collapsed/hidden rows. This example shows hows to add secondary groups
    /// within a primary grouping. Excel requires at least one row between each
    /// outline grouping at the same level.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_rows_collapsed2.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_with_format(0, 0, "Region", &bold)?;
    /// #     worksheet.write(1, 0, "North 1")?;
    /// #     worksheet.write(2, 0, "North 2")?;
    /// #     worksheet.write(3, 0, "North 3")?;
    /// #     worksheet.write(4, 0, "North 4")?;
    /// #     worksheet.write_with_format(5, 0, "North Total", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 1, "Sales", &bold)?;
    /// #     worksheet.write(1, 1, 1000)?;
    /// #     worksheet.write(2, 1, 1200)?;
    /// #     worksheet.write(3, 1, 900)?;
    /// #     worksheet.write(4, 1, 1200)?;
    /// #     worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", &bold)?;
    /// #
    /// #     worksheet.write(6, 0, "South 1")?;
    /// #     worksheet.write(7, 0, "South 2")?;
    /// #     worksheet.write(8, 0, "South 3")?;
    /// #     worksheet.write(9, 0, "South 4")?;
    /// #     worksheet.write_with_format(10, 0, "South Total", &bold)?;
    /// #
    /// #     worksheet.write(6, 1, 400)?;
    /// #     worksheet.write(7, 1, 600)?;
    /// #     worksheet.write(8, 1, 500)?;
    /// #     worksheet.write(9, 1, 600)?;
    /// #     worksheet.write_formula_with_format(10, 1, "=SUBTOTAL(9,B7:B10)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(11, 0, "Grand Total", &bold)?;
    /// #     worksheet.write_formula_with_format(11, 1, "=SUBTOTAL(9,B2:B11)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add grouping over the sub-total range.
    ///     worksheet.group_rows(1, 10)?;
    ///
    ///     // Add secondary collapsed groups within the first range.
    ///     worksheet.group_rows_collapsed(1, 4)?;
    ///     worksheet.group_rows_collapsed(6, 9)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_collapsed2.png">
    ///
    /// For comparison here is the expanded output:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows2.png">
    ///
    pub fn group_rows_collapsed(
        &mut self,
        first_row: RowNum,
        last_row: RowNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Set the collapsed property to the row after the collapsed group.
        self.set_collapsed_row(last_row + 1)?;

        self.set_grouped_rows(first_row, last_row, true)
    }

    /// Group a range of columns into a worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// See [`Worksheet::group_rows()`] above for an explanation on how to create
    /// sub-groupings.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column greater than the last
    ///   column. Note, to reverse the group direction see the
    ///   [`Worksheet::group_symbols_to_left()`] method.
    /// - [`XlsxError::MaxGroupLevelExceeded`] - Group depth level exceeds
    ///   Excel's limit of 8 levels.
    ///
    ///
    /// # Examples
    ///
    /// An example of how to group worksheet columns into outlines.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_columns1.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let data = [50, 20, 15, 25, 65, 80];
    ///     let headings = ["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"];
    ///
    ///     worksheet.write_row_with_format(0, 0, headings, &bold)?;
    ///     worksheet.write_row(1, 1, data)?;
    ///     worksheet.write_formula_with_format(1, 7, "=SUBTOTAL(9,B2:G2)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    ///
    ///     // Add grouping for the over the sub-total range.
    ///     worksheet.group_columns(1, 6)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_group_columns1.png">
    ///
    /// An example of how to group worksheet columns into outlines. This example
    /// shows hows to add secondary groups within a primary grouping. Excel requires
    /// at least one column between each outline grouping at the same level.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_columns2.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_with_format(0, 0, "Region", &bold)?;
    /// #     worksheet.write_with_format(1, 0, "North", &bold)?;
    /// #     worksheet.write_with_format(2, 0, "South", &bold)?;
    /// #     worksheet.write_with_format(3, 0, "East", &bold)?;
    /// #     worksheet.write_with_format(4, 0, "West", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 1, "Jan", &bold)?;
    /// #     worksheet.write(1, 1, 50)?;
    /// #     worksheet.write(2, 1, 10)?;
    /// #     worksheet.write(3, 1, 45)?;
    /// #     worksheet.write(4, 1, 15)?;
    /// #
    /// #     worksheet.write_with_format(0, 2, "Feb", &bold)?;
    /// #     worksheet.write(1, 2, 20)?;
    /// #     worksheet.write(2, 2, 20)?;
    /// #     worksheet.write(3, 2, 75)?;
    /// #     worksheet.write(4, 2, 15)?;
    /// #
    /// #     worksheet.write_with_format(0, 3, "Mar", &bold)?;
    /// #     worksheet.write(1, 3, 15)?;
    /// #     worksheet.write(2, 3, 30)?;
    /// #     worksheet.write(3, 3, 50)?;
    /// #     worksheet.write(4, 3, 35)?;
    /// #
    /// #     worksheet.write_with_format(0, 4, "Q1 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 4, "=SUBTOTAL(9,B2:D2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 4, "=SUBTOTAL(9,B3:D3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 4, "=SUBTOTAL(9,B4:D4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 4, "=SUBTOTAL(9,B5:D5)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 5, "Apr", &bold)?;
    /// #     worksheet.write(1, 5, 25)?;
    /// #     worksheet.write(2, 5, 50)?;
    /// #     worksheet.write(3, 5, 15)?;
    /// #     worksheet.write(4, 5, 35)?;
    /// #
    /// #     worksheet.write_with_format(0, 6, "May", &bold)?;
    /// #     worksheet.write(1, 6, 65)?;
    /// #     worksheet.write(2, 6, 50)?;
    /// #     worksheet.write(3, 6, 75)?;
    /// #     worksheet.write(4, 6, 70)?;
    /// #
    /// #     worksheet.write_with_format(0, 7, "Jun", &bold)?;
    /// #     worksheet.write(1, 7, 80)?;
    /// #     worksheet.write(2, 7, 50)?;
    /// #     worksheet.write(3, 7, 90)?;
    /// #     worksheet.write(4, 7, 50)?;
    /// #
    /// #     worksheet.write_with_format(0, 8, "Q2 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 8, "=SUBTOTAL(9,F2:H2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 8, "=SUBTOTAL(9,F3:H3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 8, "=SUBTOTAL(9,F4:H4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 8, "=SUBTOTAL(9,F5:H5)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 9, "H1 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 9, "=SUBTOTAL(9,B2:I2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 9, "=SUBTOTAL(9,B3:I3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 9, "=SUBTOTAL(9,B4:I4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 9, "=SUBTOTAL(9,B5:I5)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add grouping for the over the sub-total range.
    ///     worksheet.group_columns(1, 8)?;
    ///
    ///     // Add secondary groups within the first range.
    ///     worksheet.group_columns(1, 3)?;
    ///     worksheet.group_columns(5, 7)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_group_columns2.png">
    ///
    pub fn group_columns(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.set_grouped_columns(first_col, last_col, false)
    }

    /// Group a range of columns into a collapsed worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// See [`Worksheet::group_rows()`] above for an explanation on how to
    /// create sub-groupings.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column greater than the
    ///   last column. Note, to reverse the group direction see the
    ///   [`Worksheet::group_symbols_to_left()`] method.
    /// - [`XlsxError::MaxGroupLevelExceeded`] - Group depth level exceeds
    ///   Excel's limit of 7 levels.
    ///
    ///
    /// # Examples
    ///
    /// An example of how to group worksheet columns into outlines with
    /// collapsed/hidden rows.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_columns_collapsed1.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let data = [50, 20, 15, 25, 65, 80];
    ///     let headings = ["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"];
    ///
    ///     worksheet.write_row_with_format(0, 0, headings, &bold)?;
    ///     worksheet.write_row(1, 1, data)?;
    ///     worksheet.write_formula_with_format(1, 7, "=SUBTOTAL(9,B2:G2)", &bold)?;
    ///
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add collapse grouping over the sub-total range.
    ///     worksheet.group_columns_collapsed(1, 6)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_columns_collapsed1.png">
    ///
    /// For comparison here is the expanded output:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_columns1.png">
    ///
    ///
    /// An example of how to group worksheet columns into outlines with
    /// collapsed/hidden rows. This example shows hows to add secondary groups
    /// within a primary grouping. Excel requires at least one column between
    /// each outline grouping at the same level.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_columns_collapsed2.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_with_format(0, 0, "Region", &bold)?;
    /// #     worksheet.write_with_format(1, 0, "North", &bold)?;
    /// #     worksheet.write_with_format(2, 0, "South", &bold)?;
    /// #     worksheet.write_with_format(3, 0, "East", &bold)?;
    /// #     worksheet.write_with_format(4, 0, "West", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 1, "Jan", &bold)?;
    /// #     worksheet.write(1, 1, 50)?;
    /// #     worksheet.write(2, 1, 10)?;
    /// #     worksheet.write(3, 1, 45)?;
    /// #     worksheet.write(4, 1, 15)?;
    /// #
    /// #     worksheet.write_with_format(0, 2, "Feb", &bold)?;
    /// #     worksheet.write(1, 2, 20)?;
    /// #     worksheet.write(2, 2, 20)?;
    /// #     worksheet.write(3, 2, 75)?;
    /// #     worksheet.write(4, 2, 15)?;
    /// #
    /// #     worksheet.write_with_format(0, 3, "Mar", &bold)?;
    /// #     worksheet.write(1, 3, 15)?;
    /// #     worksheet.write(2, 3, 30)?;
    /// #     worksheet.write(3, 3, 50)?;
    /// #     worksheet.write(4, 3, 35)?;
    /// #
    /// #     worksheet.write_with_format(0, 4, "Q1 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 4, "=SUBTOTAL(9,B2:D2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 4, "=SUBTOTAL(9,B3:D3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 4, "=SUBTOTAL(9,B4:D4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 4, "=SUBTOTAL(9,B5:D5)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 5, "Apr", &bold)?;
    /// #     worksheet.write(1, 5, 25)?;
    /// #     worksheet.write(2, 5, 50)?;
    /// #     worksheet.write(3, 5, 15)?;
    /// #     worksheet.write(4, 5, 35)?;
    /// #
    /// #     worksheet.write_with_format(0, 6, "May", &bold)?;
    /// #     worksheet.write(1, 6, 65)?;
    /// #     worksheet.write(2, 6, 50)?;
    /// #     worksheet.write(3, 6, 75)?;
    /// #     worksheet.write(4, 6, 70)?;
    /// #
    /// #     worksheet.write_with_format(0, 7, "Jun", &bold)?;
    /// #     worksheet.write(1, 7, 80)?;
    /// #     worksheet.write(2, 7, 50)?;
    /// #     worksheet.write(3, 7, 90)?;
    /// #     worksheet.write(4, 7, 50)?;
    /// #
    /// #     worksheet.write_with_format(0, 8, "Q2 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 8, "=SUBTOTAL(9,F2:H2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 8, "=SUBTOTAL(9,F3:H3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 8, "=SUBTOTAL(9,F4:H4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 8, "=SUBTOTAL(9,F5:H5)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 9, "H1 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 9, "=SUBTOTAL(9,B2:I2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 9, "=SUBTOTAL(9,B3:I3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 9, "=SUBTOTAL(9,B4:I4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 9, "=SUBTOTAL(9,B5:I5)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add grouping for the over the sub-total range.
    ///     worksheet.group_columns(1, 8)?;
    ///
    ///     // Add secondary collapsed groups within the first range.
    ///     worksheet.group_columns_collapsed(1, 3)?;
    ///     worksheet.group_columns_collapsed(5, 7)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_columns_collapsed2.png">
    ///
    /// For comparison here is the expanded output:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_columns2.png">
    ///
    pub fn group_columns_collapsed(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Set the collapsed property to the column after the collapsed group.
        self.set_collapsed_column(last_col + 1)?;

        self.set_grouped_columns(first_col, last_col, true)
    }

    /// Place the row outline group expand/collapse symbols above the range.
    ///
    /// This method toggles the Excel worksheet option to place the outline
    /// group expand/collapse symbols `[+]` and `[-]` above the group ranges
    /// instead of below for row ranges.
    ///
    /// In Excel this is a worksheet wide option and will apply to all row
    /// outlines in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of how to group worksheet rows into outlines. This example
    /// puts the expand/collapse symbol above the range for all row groups in
    /// the worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_symbols_above.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_with_format(0, 0, "Region", &bold)?;
    /// #     worksheet.write(1, 0, "North 1")?;
    /// #     worksheet.write(2, 0, "North 2")?;
    /// #     worksheet.write(3, 0, "North 3")?;
    /// #     worksheet.write(4, 0, "North 4")?;
    /// #     worksheet.write_with_format(5, 0, "North Total", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 1, "Sales", &bold)?;
    /// #     worksheet.write(1, 1, 1000)?;
    /// #     worksheet.write(2, 1, 1200)?;
    /// #     worksheet.write(3, 1, 900)?;
    /// #     worksheet.write(4, 1, 1200)?;
    /// #     worksheet.write_formula_with_format(5, 1, "=SUBTOTAL(9,B2:B5)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add grouping over the sub-total range.
    ///     worksheet.group_rows(1, 4)?;
    ///
    ///     // Display the expand/collapse symbol above the range.
    ///     worksheet.group_symbols_above(true);
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_symbols_above.png">
    ///
    pub fn group_symbols_above(&mut self, enable: bool) -> &mut Worksheet {
        self.outline_symbols_above = enable;
        self
    }

    /// Place the column outline group expand/collapse symbols to the left of
    /// the range.
    ///
    /// This method toggles the Excel worksheet option to place the outline
    /// group expand/collapse symbols `[+]` and `[-]` to the left of the group
    /// ranges instead of to the right, for column ranges.
    ///
    /// In Excel this is a worksheet wide option and will apply to all column
    /// outlines in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of how to group worksheet columns into outlines. This example
    /// puts the expand/collapse symbol to the left of the range for all row groups
    /// in the worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_group_symbols_to_left.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a format to use in headings.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Add a worksheet with some sample data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_with_format(0, 0, "Region", &bold)?;
    /// #     worksheet.write_with_format(1, 0, "North", &bold)?;
    /// #     worksheet.write_with_format(2, 0, "South", &bold)?;
    /// #     worksheet.write_with_format(3, 0, "East", &bold)?;
    /// #     worksheet.write_with_format(4, 0, "West", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 1, "Jan", &bold)?;
    /// #     worksheet.write(1, 1, 50)?;
    /// #     worksheet.write(2, 1, 10)?;
    /// #     worksheet.write(3, 1, 45)?;
    /// #     worksheet.write(4, 1, 15)?;
    /// #
    /// #     worksheet.write_with_format(0, 2, "Feb", &bold)?;
    /// #     worksheet.write(1, 2, 20)?;
    /// #     worksheet.write(2, 2, 20)?;
    /// #     worksheet.write(3, 2, 75)?;
    /// #     worksheet.write(4, 2, 15)?;
    /// #
    /// #     worksheet.write_with_format(0, 3, "Mar", &bold)?;
    /// #     worksheet.write(1, 3, 15)?;
    /// #     worksheet.write(2, 3, 30)?;
    /// #     worksheet.write(3, 3, 50)?;
    /// #     worksheet.write(4, 3, 35)?;
    /// #
    /// #     worksheet.write_with_format(0, 4, "Q1 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 4, "=SUBTOTAL(9,B2:D2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 4, "=SUBTOTAL(9,B3:D3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 4, "=SUBTOTAL(9,B4:D4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 4, "=SUBTOTAL(9,B5:D5)", &bold)?;
    /// #
    /// #     worksheet.write_with_format(0, 5, "Apr", &bold)?;
    /// #     worksheet.write(1, 5, 25)?;
    /// #     worksheet.write(2, 5, 50)?;
    /// #     worksheet.write(3, 5, 15)?;
    /// #     worksheet.write(4, 5, 35)?;
    /// #
    /// #     worksheet.write_with_format(0, 6, "May", &bold)?;
    /// #     worksheet.write(1, 6, 65)?;
    /// #     worksheet.write(2, 6, 50)?;
    /// #     worksheet.write(3, 6, 75)?;
    /// #     worksheet.write(4, 6, 70)?;
    /// #
    /// #     worksheet.write_with_format(0, 7, "Jun", &bold)?;
    /// #     worksheet.write(1, 7, 80)?;
    /// #     worksheet.write(2, 7, 50)?;
    /// #     worksheet.write(3, 7, 90)?;
    /// #     worksheet.write(4, 7, 50)?;
    /// #
    /// #     worksheet.write_with_format(0, 8, "Q2 Total", &bold)?;
    /// #     worksheet.write_formula_with_format(1, 8, "=SUBTOTAL(9,F2:H2)", &bold)?;
    /// #     worksheet.write_formula_with_format(2, 8, "=SUBTOTAL(9,F3:H3)", &bold)?;
    /// #     worksheet.write_formula_with_format(3, 8, "=SUBTOTAL(9,F4:H4)", &bold)?;
    /// #     worksheet.write_formula_with_format(4, 8, "=SUBTOTAL(9,F5:H5)", &bold)?;
    /// #
    /// #     // Autofit the columns for clarity.
    /// #     worksheet.autofit();
    /// #
    ///     // Add groupings for the over the sub-total ranges.
    ///     worksheet.group_columns(1, 3)?;
    ///     worksheet.group_columns(5, 7)?;
    ///
    ///     // Display the expand/collapse symbol to the left of the range.
    ///     worksheet.group_symbols_to_left(true);
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_group_symbols_to_left.png">
    ///
    ///
    pub fn group_symbols_to_left(&mut self, enable: bool) -> &mut Worksheet {
        self.outline_symbols_left = enable;
        self
    }

    /// Hide a worksheet row.
    ///
    /// The `set_row_hidden()` method is used to hide a row. This can be
    /// used, for example, to hide intermediary steps in a complicated
    /// calculation.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates hiding a worksheet row.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_row_hidden.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Hide row 2 (with zero indexing).
    ///     worksheet.set_row_hidden(1)?;
    ///
    ///     worksheet.write_string(2, 0, "Row 2 is hidden")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_row_hidden.png">
    ///
    pub fn set_row_hidden(&mut self, row: RowNum) -> Result<&mut Worksheet, XlsxError> {
        // Get a suitable minimum value for the dimension check/set.
        let min_col = self.get_min_col();

        // Check row is in the allowed range.
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing row metadata object or create a new one.
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => row_options.hidden = true,
            None => {
                let row_options = RowOptions {
                    height: self.user_default_row_height,
                    xf_index: 0,
                    hidden: true,
                    level: 0,
                    collapsed: false,
                    format: None,
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    /// Unhide a user hidden worksheet row.
    ///
    /// The `set_row_unhidden()` method is used to unhide a previously hidden
    /// row. This can occasionally be useful when used in conjunction with
    /// autofilter rules.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row exceeds Excel's worksheet
    ///   limits.
    ///
    pub fn set_row_unhidden(&mut self, row: RowNum) -> Result<&mut Worksheet, XlsxError> {
        // Get a suitable minimum value for the dimension check/set.
        let min_col = self.get_min_col();

        // Check row is in the allowed range.
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Only update an existing row metadata object.
        if let Some(row_options) = self.changed_rows.get_mut(&row) {
            row_options.hidden = false;
        }

        Ok(self)
    }

    /// Set the default row height for all rows in a worksheet, efficiently.
    ///
    /// This method can be used to efficiently set the default row height for
    /// all rows in a worksheet. It is efficient because it uses an Excel
    /// optimization to adjust the row heights with a single XML element. By
    /// contrast, using [`Worksheet::set_row_height()`] for every row in a
    /// worksheet would require a million XML elements and would result in a
    /// very large file.
    ///
    /// The height is specified in character units, where the default height is
    /// 15. Excel allows height values in increments of 0.25.
    ///
    /// Individual row heights can be set via [`Worksheet::set_row_height()`].
    ///
    /// Note, there is no equivalent method for columns because the file format
    /// already optimizes the storage of a large number of contiguous columns.
    ///
    /// # Parameters
    ///
    /// - `height`: The row height in character units. Must be greater than 0.0.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the default row height for
    /// all rows in a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_default_row_height.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set the default row height in Excel character units.
    ///     worksheet.set_default_row_height(30);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_default_row_height.png">
    ///
    pub fn set_default_row_height(&mut self, height: impl Into<f64>) -> &mut Worksheet {
        let height = height.into();
        if height <= 0.0 {
            return self;
        }

        self.user_default_row_height = height;
        self
    }

    /// Set the default row height in pixels for all rows in a worksheet,
    /// efficiently.
    ///
    /// The `set_default_row_height_pixels()` method is used to change the
    /// default height of all rows in a worksheet. See
    /// [`Worksheet::set_default_row_height()`] above for an explanation.
    ///
    /// The height is specified in pixels, where the default height is 20.
    ///
    /// # Parameters
    ///
    /// - `height`: The row height in pixels.
    ///
    pub fn set_default_row_height_pixels(&mut self, height: u16) -> &mut Worksheet {
        let height = 0.75 * f64::from(height);
        self.set_default_row_height(height)
    }

    /// Hide all unused rows in a worksheet, efficiently.
    ///
    /// This method can be used to efficiently hide unused rows in a worksheet.
    /// It is efficient because it uses an Excel optimization to hide the rows
    /// with a single XML element. By contrast, using
    /// [`Worksheet::set_row_hidden()`] for the majority of rows in a worksheet
    /// would require a million XML elements and would result in a very large
    /// file.
    ///
    /// "Unused" in this context means that the row doesn't contain data,
    /// formatting, or any changes such as the row height.
    ///
    /// To hide individual rows use the [`Worksheet::set_row_hidden()`] method.
    ///
    /// Note, there is no equivalent method for columns because the file format
    /// already optimizes the storage of a large number of contiguous columns.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates efficiently hiding the unused rows in
    /// a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_hide_unused_rows.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write some data.
    ///     worksheet.write(0, 0, "First row")?;
    ///     worksheet.write(6, 0, "Last row")?;
    ///
    ///     // Set the row height for the blank rows so that they are "used".
    ///     for row in 1..6 {
    ///         worksheet.set_row_height(row, 15)?;
    ///     }
    ///
    ///     // Hide all the unused rows.
    ///     worksheet.hide_unused_rows(true);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_hide_unused_rows.png">
    ///
    pub fn hide_unused_rows(&mut self, enable: bool) -> &mut Worksheet {
        self.hide_unused_rows = enable;
        self
    }

    /// Set the width for a worksheet column.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// worksheet column.
    ///
    /// The `width` parameter sets the column width in the same units used by
    /// Excel which is: the number of characters in the default font. The
    /// default width is 8.43 in the default font of Calibri 11. The actual
    /// relationship between a string width and a column width in Excel is
    /// complex. See the [following explanation of column
    /// widths](https://support.microsoft.com/en-us/kb/214123) from the
    /// Microsoft support documentation for more details. To set the width in
    /// pixels use the [`Worksheet::set_column_width_pixels()`] method.
    ///
    /// See also the [`Worksheet::autofit()`] method.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `width`: The column width in character units.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the width of columns in
    /// Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_width.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string(0, 0, "Normal")?;
    ///     worksheet.write_string(0, 2, "Wider")?;
    ///     worksheet.write_string(0, 4, "Narrower")?;
    ///
    ///     // Set the column width in Excel character units.
    ///     worksheet.set_column_width(2, 16)?;
    ///     worksheet.set_column_width(4, 4)?;
    ///     worksheet.set_column_width(5, 4)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_column_width.png">
    ///
    pub fn set_column_width(
        &mut self,
        col: ColNum,
        width: impl Into<f64>,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.set_column_width_internal(col, width, false)
    }

    /// Set the width for a worksheet column in pixels.
    ///
    /// The `set_column_width_pixels()` method is used to change the default
    /// width in pixels of a worksheet column.
    ///
    /// To set the width in Excel character units use the
    /// [`Worksheet::set_column_width()`] method.
    ///
    /// See also the [`Worksheet::autofit()`] method.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `width`: The column width in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the width of columns in Excel
    /// in pixels.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_width_pixels.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some text.
    ///     worksheet.write_string(0, 0, "Normal")?;
    ///     worksheet.write_string(0, 2, "Wider")?;
    ///     worksheet.write_string(0, 4, "Narrower")?;
    ///
    ///     // Set the column width in pixels.
    ///     worksheet.set_column_width_pixels(2, 117)?;
    ///     worksheet.set_column_width_pixels(4, 33)?;
    ///     worksheet.set_column_width_pixels(5, 33)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_column_width.png">
    ///
    pub fn set_column_width_pixels(
        &mut self,
        col: ColNum,
        width: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.set_column_width_internal(col, Self::pixels_to_width(width), false)
    }

    /// Set the column pixel width in auto-fit mode.
    ///
    /// In Excel the width of an auto-fitted column will increase if the user
    /// edits a number and increases the number of digits past the previous
    /// maximum width. This behavior doesn't apply to strings or when the number
    /// of digits is decreased. It also doesn't apply to columns that have been
    /// set manually.
    ///
    /// The `Worksheet::set_column_autofit_width()` method emulates this auto-fit
    /// behavior whereas the [`Worksheet::set_column_width_pixels()`] method,
    /// see above, is equivalent to setting the width manually.
    ///
    /// The distinction is subtle and most users are unaware of this behavior in
    /// Excel. However, it is supported for users who wish to implement their
    /// own version of auto-fit.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `width`: The column width in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    pub fn set_column_autofit_width(
        &mut self,
        col: ColNum,
        width: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.set_column_width_internal(col, Self::pixels_to_width(width), true)
    }

    /// Set the format for a column of cells.
    ///
    /// The `set_column_format()` method is used to change the default format of a
    /// column. Any unformatted data written to that column will then adopt that
    /// format. Formatted data written to the column will maintain its own cell
    /// format. See the example below.
    ///
    /// A future version of this library may support automatic merging of
    /// explicit cell formatting with the column formatting but that isn't
    /// currently supported.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format for a column in Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add for formats.
    ///     let bold_format = Format::new().set_bold();
    ///     let red_format = Format::new().set_font_color(Color::Red);
    ///
    ///     // Set the column format.
    ///     worksheet.set_column_format(1, &red_format)?;
    ///
    ///     // Add some unformatted text that adopts the column format.
    ///     worksheet.write_string(0, 1, "Hello")?;
    ///
    ///     // Add some formatted text that overrides the column format.
    ///     worksheet.write_string_with_format(2, 1, "Hello", &bold_format)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_column_format.png">
    ///
    pub fn set_column_format(
        &mut self,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Get a suitable minimum value for the dimension check/set.
        let min_row = self.get_min_row();

        // Check column is in the allowed range.
        if !self.check_dimensions(min_row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_xf_index(format);

        // Update an existing col metadata object or create a new one.
        match self.changed_cols.get_mut(&col) {
            Some(col_options) => col_options.xf_index = xf_index,
            None => {
                let col_options = ColOptions {
                    width: DEFAULT_COL_WIDTH,
                    xf_index,
                    hidden: false,
                    level: 0,
                    collapsed: false,
                    autofit: false,
                    format: Some(format.clone()),
                };
                self.changed_cols.insert(col, col_options);
            }
        }

        Ok(self)
    }

    /// Hide a worksheet column.
    ///
    /// The `set_column_hidden()` method is used to hide a column. This can be
    /// used, for example, to hide intermediary steps in a complicated
    /// calculation.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates hiding a worksheet column.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_hidden.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Hide column B.
    ///     worksheet.set_column_hidden(1)?;
    ///
    ///     worksheet.write_string(0, 3, "Column B is hidden")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_column_hidden.png">
    ///
    pub fn set_column_hidden(&mut self, col: ColNum) -> Result<&mut Worksheet, XlsxError> {
        // Check if column is in the allowed range without updating dimensions.
        if col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing col metadata object or create a new one.
        match self.changed_cols.get_mut(&col) {
            Some(col_options) => col_options.hidden = true,
            None => {
                let col_options = ColOptions {
                    width: DEFAULT_COL_WIDTH,
                    xf_index: 0,
                    hidden: true,
                    level: 0,
                    collapsed: false,
                    autofit: false,
                    format: None,
                };
                self.changed_cols.insert(col, col_options);
            }
        }

        Ok(self)
    }

    /// Set the width for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the width for a range of
    /// contiguous cells. See [`Worksheet::set_column_width()`] for more
    /// details on the single column version.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    /// - `width`: The column width in character units.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column larger than the last
    ///   column.
    ///
    pub fn set_column_range_width(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        width: impl Into<f64>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check order of first/last values.
        if first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        let width = width.into();
        for col_num in first_col..=last_col {
            self.set_column_width(col_num, width)?;
        }

        Ok(self)
    }

    /// Set the pixel width for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the width in pixels for a range of
    /// contiguous cells. See [`Worksheet::set_column_width_pixels()`] for more
    /// details on the single column version.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    /// - `width`: The column width in pixels.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column larger than the last
    ///   column.
    ///
    pub fn set_column_range_width_pixels(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        width: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check order of first/last values.
        if first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        for col_num in first_col..=last_col {
            self.set_column_width_pixels(col_num, width)?;
        }

        Ok(self)
    }

    /// Set the format for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the format for a range of
    /// contiguous cells. See [`Worksheet::set_column_format()`] for more
    /// details on the single column version.
    ///
    /// Note, this method can be used to set the cell format for the entire
    /// worksheet when applied to all the columns in the worksheet (see the
    /// example below). This is relatively efficient since it is stored as a
    /// single XML element. This is also how Excel applies a format to an entire
    /// worksheet.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column larger than the last
    ///   column.
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format for all the
    /// columns in an Excel worksheet. This effectively, and efficiently, sets
    /// the format for the entire worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_column_range_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add a format.
    ///     let cell_format = Format::new()
    ///         .set_background_color("#F9F2EC")
    ///         .set_border(FormatBorder::Thin);
    ///
    ///     // Set the column format for the entire worksheet.
    ///     worksheet.set_column_range_format(0, 16_383, &cell_format)?;
    ///
    ///     // Add some unformatted text that adopts the column format.
    ///     worksheet.write_string(1, 1, "Hello")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_column_range_format.png">
    ///
    pub fn set_column_range_format(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check order of first/last values.
        if first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        for col_num in first_col..=last_col {
            self.set_column_format(col_num, format)?;
        }

        Ok(self)
    }

    /// Hide a range of worksheet columns.
    ///
    /// This is a syntactic shortcut for hiding a range of contiguous cells. See
    /// [`Worksheet::set_column_hidden()`] for more details on the single column
    /// version.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column larger than the last
    ///   column.
    ///
    pub fn set_column_range_hidden(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check order of first/last values.
        if first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        for col_num in first_col..=last_col {
            self.set_column_hidden(col_num)?;
        }

        Ok(self)
    }

    /// Set the autofilter area in the worksheet.
    ///
    /// The `autofilter()` method allows an autofilter to be added to a
    /// worksheet. An autofilter is a way of adding drop down lists to the
    /// headers of a 2D range of worksheet data. This allows users to filter the
    /// data based on simple criteria so that some data is shown and some is
    /// hidden.
    ///
    /// See the [`Worksheet::filter_column`] method for an explanation of how to
    /// set a filter conditions for columns in the autofilter range.
    ///
    /// Note, Excel only allows one autofilter range per worksheet so calling
    /// this method multiple times will overwrite the previous range.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    /// - [`XlsxError::AutofilterRangeOverlaps`] - The autofilter range overlaps
    ///   a table autofilter range.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting a simple autofilter in a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_autofilter.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with some sample data to filter.
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(0, 0, "Region")?;
    /// #     worksheet.write_string(1, 0, "East")?;
    /// #     worksheet.write_string(2, 0, "West")?;
    /// #     worksheet.write_string(3, 0, "East")?;
    /// #     worksheet.write_string(4, 0, "North")?;
    /// #     worksheet.write_string(5, 0, "South")?;
    /// #     worksheet.write_string(6, 0, "West")?;
    /// #
    /// #     worksheet.write_string(0, 1, "Sales")?;
    /// #     worksheet.write_number(1, 1, 3000)?;
    /// #     worksheet.write_number(2, 1, 8000)?;
    /// #     worksheet.write_number(3, 1, 5000)?;
    /// #     worksheet.write_number(4, 1, 4000)?;
    /// #     worksheet.write_number(5, 1, 7000)?;
    /// #     worksheet.write_number(6, 1, 9000)?;
    /// #
    ///     // Set the autofilter.
    ///     worksheet.autofilter(0, 0, 6, 1)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_autofilter.png">
    ///
    pub fn autofilter(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Store the defined name information.
        self.autofilter_defined_name.in_use = true;
        self.autofilter_defined_name.name_type = DefinedNameType::Autofilter;
        self.autofilter_defined_name.first_row = first_row;
        self.autofilter_defined_name.first_col = first_col;
        self.autofilter_defined_name.last_row = last_row;
        self.autofilter_defined_name.last_col = last_col;

        let autofilter_area = CellRange::new(first_row, first_col, last_row, last_col);

        // Clear any previous filters.
        self.filter_conditions = BTreeMap::new();

        // Store the cells with the autofilter dropdown for the autofit calc.
        for col in first_col..=last_col {
            // Check that the worksheet autofilter doesn't overlap a table
            // autofilter.
            let filter_data = self.cells_with_autofilter.get(&(first_row, col));
            if let Some((filter_type, filter_range)) = filter_data {
                if *filter_type == FilterType::Table {
                    return Err(XlsxError::AutofilterRangeOverlaps(
                        autofilter_area.to_error_string(),
                        filter_range.to_error_string(),
                    ));
                }
            }

            self.cells_with_autofilter.insert(
                (first_row, col),
                (FilterType::Worksheet, autofilter_area.clone()),
            );
        }

        self.autofilter_area = autofilter_area.to_range_string();

        Ok(self)
    }

    /// Set the filter condition for a column in an autofilter range.
    ///
    /// The [`Worksheet::autofilter()`] method sets the cell range for an
    /// autofilter but in order to filter rows within the filter area you must
    /// also add a filter condition.
    ///
    /// Excel supports two main types of filter. The first, and most common, is
    /// a list filter where the user selects the items to filter from a list of
    /// all the values in the the column range:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/autofilter_list.png">
    ///
    /// The other main type of filter is a custom filter where the user can
    /// specify 1 or 2 conditions like ">= 4000" and "<= 6000":
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/autofilter_custom.png">
    ///
    /// In Excel these are mutually exclusive and you will need to choose one or
    /// the other via the [`FilterCondition`] struct parameter.
    ///
    /// For more details on setting filter conditions see [`FilterCondition`]
    /// and the [Working with Autofilters] section of the Users Guide.
    ///
    /// [Working with Autofilters]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html
    ///
    /// Note, there are some limitations on autofilter conditions. The main one
    /// is that the hiding of rows that don't match a filter is not an automatic
    /// part of the file format. Instead it is necessary to hide rows that don't
    /// match the filters. The `rust_xlsxwriter` library does this automatically
    /// and in most cases will get it right, however, there may be cases where
    /// you need to manually hide some of the rows. See [Auto-hiding filtered
    /// rows].
    ///
    /// [Auto-hiding filtered rows]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html#auto-hiding-filtered-rows
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `filter_condition`: The column filter condition defined by the
    ///   [`FilterCondition`] struct.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::ParameterError`] - Parameter error for the following
    ///   issues:
    ///   - The [`Worksheet::autofilter()`] range hasn't been set.
    ///   - The column is outside the [`Worksheet::autofilter()`] range.
    ///   - The [`FilterCondition`] doesn't have a condition set.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting an autofilter with a list
    /// filter condition.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_filter_column1.rs
    /// #
    /// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with some sample data to filter.
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(0, 0, "Region")?;
    /// #     worksheet.write_string(1, 0, "East")?;
    /// #     worksheet.write_string(2, 0, "West")?;
    /// #     worksheet.write_string(3, 0, "East")?;
    /// #     worksheet.write_string(4, 0, "North")?;
    /// #     worksheet.write_string(5, 0, "South")?;
    /// #     worksheet.write_string(6, 0, "West")?;
    /// #
    /// #     worksheet.write_string(0, 1, "Sales")?;
    /// #     worksheet.write_number(1, 1, 3000)?;
    /// #     worksheet.write_number(2, 1, 8000)?;
    /// #     worksheet.write_number(3, 1, 5000)?;
    /// #     worksheet.write_number(4, 1, 4000)?;
    /// #     worksheet.write_number(5, 1, 7000)?;
    /// #     worksheet.write_number(6, 1, 9000)?;
    /// #
    /// #     // Set the autofilter.
    /// #     worksheet.autofilter(0, 0, 6, 1)?;
    /// #
    ///     // Set a filter condition to only show cells matching "East" in the first
    ///     // column.
    ///     let filter_condition = FilterCondition::new().add_list_filter("East");
    ///     worksheet.filter_column(0, &filter_condition)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column1.png">
    ///
    pub fn filter_column(
        &mut self,
        col: ColNum,
        filter_condition: &FilterCondition,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check if column is in the allowed range without updating dimensions.
        if col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check that an autofilter has been created before a condition can be
        // applied to it.
        if !self.autofilter_defined_name.in_use {
            let error =
                "The 'autofilter()' range must be set before a 'filter_condition' can be applied."
                    .to_string();
            return Err(XlsxError::ParameterError(error));
        }

        // Check if column is within the autofilter column range.
        if col < self.autofilter_defined_name.first_col
            || col > self.autofilter_defined_name.last_col
        {
            let error = format!(
                "Col '{col}' outside user defined autofilter column range '{}-{}'",
                self.autofilter_defined_name.first_col, self.autofilter_defined_name.last_col
            );
            return Err(XlsxError::ParameterError(error));
        }

        // Check the filter condition have been set up correctly.
        if filter_condition.list.is_empty()
            && filter_condition.custom1.is_none()
            && !filter_condition.should_match_blanks
        {
            let error =
                "The 'filter_condition' doesn't have a data value or condition set.".to_string();
            return Err(XlsxError::ParameterError(error));
        }

        self.filter_conditions.insert(col, filter_condition.clone());

        Ok(self)
    }

    /// Turn off the option to automatically hide rows that don't match filters.
    ///
    /// Rows that don't match autofilter conditions are hidden by Excel at
    /// runtime. This feature isn't an automatic part of the file format and in
    /// practice it is necessary for the user to hide rows that don't match the
    /// applied filters. The `rust_xlsxwriter` library tries to do this
    /// automatically and in most cases will get it right, however, there may be
    /// cases where you need to manually hide some of the rows and may want to
    /// turn off the automatic handling using `filter_automatic_off()`.
    ///
    /// See [Auto-hiding filtered rows] in the User Guide.
    ///
    /// [Auto-hiding filtered rows]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html#auto-hiding-filtered-rows
    ///
    pub fn filter_automatic_off(&mut self) -> &mut Worksheet {
        self.filter_automatic_off = true;
        self
    }

    /// Add a table to a worksheet.
    ///
    /// Tables in Excel are a way of grouping a range of cells into a single
    /// entity that has common formatting or that can be referenced from
    /// formulas. Tables can have column headers, autofilters, total rows,
    /// column formulas and different formatting styles.
    ///
    /// The headers and total row of a table should be configured via a
    /// [`Table`] struct but the table data can be added via standard
    /// [`Worksheet::write()`] methods.
    ///
    /// For more information on tables see the Microsoft documentation on
    /// [Overview of Excel tables].
    ///
    /// [Overview of Excel tables]:
    ///     https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// Note, you need to ensure that the `first_row` and `last_row` range
    /// includes all the rows for the table including the header and the total
    /// row, if present.
    ///
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    /// - [`XlsxError::TableError`] - A general error that is raised when a
    ///   table parameter is incorrect or a table is configured incorrectly.
    /// - [`XlsxError::TableRangeOverlaps`] - The table range overlaps a
    ///   previous table range.
    /// - [`XlsxError::AutofilterRangeOverlaps`] - The table autofilter range
    ///   overlaps the worksheet autofilter range.
    ///
    /// # Examples
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_columns.rs
    /// #
    /// use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some sample data for the table.
    ///     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    ///     let data = [
    ///         [10000, 5000, 8000, 6000],
    ///         [2000, 3000, 4000, 5000],
    ///         [6000, 6000, 6500, 6000],
    ///         [500, 300, 200, 700],
    ///     ];
    ///
    ///     // Write the table data.
    ///     worksheet.write_column(3, 1, items)?;
    ///     worksheet.write_row_matrix(3, 2, data)?;
    ///
    ///     // Set the column widths for clarity.
    ///     worksheet.set_column_range_width(1, 6, 12)?;
    ///
    ///     // Create a new table and configure it.
    ///     let columns = vec![
    ///         TableColumn::new()
    ///             .set_header("Product")
    ///             .set_total_label("Totals"),
    ///         TableColumn::new()
    ///             .set_header("Quarter 1")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Quarter 2")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Quarter 3")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Quarter 4")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Year")
    ///             .set_total_function(TableFunction::Sum)
    ///             .set_formula("SUM(Table1[@[Quarter 1]:[Quarter 4]])"),
    ///     ];
    ///
    ///     let table = Table::new().set_columns(&columns).set_total_row(true);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 7, 6, &table)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("tables.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_columns.png">
    ///
    pub fn add_table(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        table: &Table,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Check that the table has at least one valid row.
        if first_row == last_row && table.show_header_row {
            return Err(XlsxError::TableError(
                "Table must have at least one row".to_string(),
            ));
        }

        let default_headers =
            self.default_table_headers(first_row, first_col, last_col, table.show_header_row);

        let mut table = table.clone();
        table.cell_range = CellRange::new(first_row, first_col, last_row, last_col);
        table.initialize_columns(&default_headers)?;

        let first_data_row = table.first_data_row();
        let last_data_row = table.last_data_row();

        // Write the worksheet information required for each column.
        for (offset, column) in table.columns.iter_mut().enumerate() {
            let col = first_col + offset as u16;

            // Write the header. We skip this when writing serde headers.
            if table.show_header_row && !table.is_serde_table {
                match &column.header_format {
                    Some(header_format) => {
                        self.write_string_with_format(first_row, col, &column.name, header_format)?;
                    }
                    None => {
                        self.write_string(first_row, col, &column.name)?;
                    }
                }
            }

            // In constant memory mode add any formulas or totals to the
            // write-ahead buffer.
            self.set_writing_ahead(true);

            // Get a copy of the column format or use the default format. This
            // is mainly to work around constant memory cases which can't use the
            // update_cell_format() approach below.
            let col_format = match &column.format {
                Some(format) => format.clone(),
                None => Format::default(),
            };

            // Write the total row strings or formulas.
            if table.show_total_row {
                if !column.total_label.is_empty() {
                    self.write_string_with_format(last_row, col, &column.total_label, &col_format)?;
                } else if column.total_function != TableFunction::None {
                    let formula = column.total_function();
                    self.write_formula_with_format(last_row, col, formula, &col_format)?;
                }
            }

            // Write the column formula as worksheet formulas.
            if let Some(formula) = &column.formula {
                for row in first_data_row..=last_data_row {
                    self.write_formula_with_format(row, col, formula, &col_format)?;
                }
            }

            // Set the column format local index if required.
            if let Some(format) = column.format.as_mut() {
                format.dxf_index = self.format_dxf_index(format);
                let format_index = self.format_xf_index(format);
                for row in first_data_row..=last_data_row {
                    self.update_cell_format(row, col, format_index);
                }

                if table.show_total_row && column.total_function != TableFunction::None {
                    self.update_cell_format(last_row, col, format_index);
                }
            }

            // Stop writing ahead in constant memory mode.
            self.set_writing_ahead(false);
        }

        // Create a cell range for storage and range testing.
        let cell_range = CellRange::new(first_row, first_col, last_row, last_col);

        // Check if the table range overlaps any previous table range. This is a
        // major error in Excel. Note, the ranges are stored in a separate Vec
        // to the cells to cut down on storage size.
        let new_index = self.table_ranges.len();
        for row in first_row..=last_row {
            for col in first_col..=last_col {
                match self.table_cells.get_mut(&(row, col)) {
                    Some(index) => {
                        let previous_cell_range = self.table_ranges.get(*index).unwrap();
                        return Err(XlsxError::TableRangeOverlaps(
                            cell_range.to_error_string(),
                            previous_cell_range.to_error_string(),
                        ));
                    }
                    None => self.table_cells.insert((row, col), new_index),
                };
            }
        }

        // Store the cells with the autofilter dropdown for the autofit calc.
        if table.show_autofilter {
            for col in first_col..=last_col {
                // Check that the table autofilter doesn't overlap the worksheet
                // autofilter.
                let filter_data = self.cells_with_autofilter.get(&(first_row, col));
                if let Some((filter_type, filter_range)) = filter_data {
                    if *filter_type == FilterType::Worksheet {
                        return Err(XlsxError::AutofilterRangeOverlaps(
                            table.cell_range.to_error_string(),
                            filter_range.to_error_string(),
                        ));
                    }
                }

                self.cells_with_autofilter.insert(
                    (first_row, col),
                    (FilterType::Table, table.cell_range.clone()),
                );
            }
        }

        // Store the table if everything was okay.
        self.table_ranges.push(cell_range);
        self.tables.push(table);

        Ok(self)
    }

    /// Add a conditional format to highlight cells based on rules.
    ///
    /// Conditional formatting is a feature of Excel which allows you to apply a
    /// format to a cell or a range of cells based on certain criteria. This is
    /// generally used to highlight particular values in a range of data.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/conditional_format_cell_intro.png">
    ///
    /// The [`ConditionalFormat`](crate::conditional_format) variants are used
    /// to represent the types of conditional format that can be applied in
    /// Excel.
    ///
    /// Multiple formatting rules that affect the same range or overlapping
    /// ranges are applied in the order that they are added to the worksheet.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `conditional_format`: A conditional format instance that implements
    ///   the [`ConditionalFormat`] trait.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    /// - [`XlsxError::ConditionalFormatError`] - A general error that is raised
    ///   when a conditional formatting parameter is incorrect or missing.
    ///
    /// # Examples
    ///
    /// Example of adding a cell type conditional formatting to a worksheet.
    /// Cells with values >= 50 are in light red. Values < 50 are in light
    /// green.
    ///
    /// ```
    /// # // This code is available in examples/doc_conditional_format_cell1.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some sample data.
    /// #     let data = [
    /// #         [90, 80, 50, 10, 20, 90, 40, 90, 30, 40],
    /// #         [20, 10, 90, 100, 30, 60, 70, 60, 50, 90],
    /// #         [10, 50, 60, 50, 20, 50, 80, 30, 40, 60],
    /// #         [10, 90, 20, 40, 10, 40, 50, 70, 90, 50],
    /// #         [70, 100, 10, 90, 10, 10, 20, 100, 100, 40],
    /// #         [20, 60, 10, 100, 30, 10, 20, 60, 100, 10],
    /// #         [10, 60, 10, 80, 100, 80, 30, 30, 70, 40],
    /// #         [30, 90, 60, 10, 10, 100, 40, 40, 30, 40],
    /// #         [80, 90, 10, 20, 20, 50, 80, 20, 60, 90],
    /// #         [60, 80, 30, 30, 10, 50, 80, 60, 50, 30],
    /// #     ];
    /// #     worksheet.write_row_matrix(2, 1, data)?;
    /// #
    /// #     // Set the column widths for clarity.
    /// #     worksheet.set_column_range_width(1, 10, 6)?;
    /// #
    /// #     // Add a format. Light red fill with dark red text.
    /// #     let format1 = Format::new()
    /// #         .set_font_color("9C0006")
    /// #         .set_background_color("FFC7CE");
    /// #
    /// #     // Add a format. Green fill with dark green text.
    /// #     let format2 = Format::new()
    /// #         .set_font_color("006100")
    /// #         .set_background_color("C6EFCE");
    /// #
    ///     // Write a conditional format over a range.
    ///     let conditional_format = ConditionalFormatCell::new()
    ///         .set_rule(ConditionalFormatCellRule::GreaterThanOrEqualTo(50))
    ///         .set_format(format1);
    ///
    ///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
    ///
    ///     // Write another conditional format over the same range.
    ///     let conditional_format = ConditionalFormatCell::new()
    ///         .set_rule(ConditionalFormatCellRule::LessThan(50))
    ///         .set_format(format2);
    ///
    ///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("conditional_format.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/conditional_format_cell1.png">
    ///
    pub fn add_conditional_format<T>(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        conditional_format: &T,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: ConditionalFormat + Send + Sync,
    {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        let mut conditional_format = conditional_format.box_clone();

        // Store the conditional formats based on their range.
        let mut cell_range = utility::cell_range(first_row, first_col, last_row, last_col);
        let multi_range = conditional_format.multi_range();
        if !multi_range.is_empty() {
            cell_range = multi_range;
        }

        // Validate the conditional format.
        conditional_format.validate()?;

        // Check for extended Excel 2010 data bars/icons.
        if conditional_format.has_x14_extensions() {
            self.use_x14_extensions = true;
            self.has_x14_conditional_formats = true;
        }

        // Only write standard cond formats for non-x14 icons.
        if !conditional_format.has_x14_only() {
            self.has_conditional_formats = true;
        }

        // Set the dxf format local index if required.
        if let Some(format) = conditional_format.format_as_mut() {
            format.dxf_index = self.format_dxf_index(format);
        }

        match self.conditional_formats.entry(cell_range.clone()) {
            Entry::Occupied(mut entry) => {
                // The conditional format range already exists. Append the rule.
                let rules = entry.get_mut();
                rules.push(conditional_format);
            }
            Entry::Vacant(entry) => {
                // The row doesn't exist, create a new row with columns and insert
                // the cell value.
                let rules = vec![conditional_format];
                entry.insert(rules);

                // Also store the order that the ranges are added in.
                self.conditional_format_order.push(cell_range);
            }
        }

        Ok(self)
    }

    /// Add a data validation to one or more cells to restrict user input based
    /// on types and rules.
    ///
    /// Data validation is a feature of Excel which allows you to restrict the
    /// data that a user enters in a cell and to display associated help and
    /// warning messages. It also allows you to restrict input to values in a
    /// dropdown list.
    ///
    /// See [`DataValidation`] for more information and examples.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `data_validation`: A [`DataValidation`] data validation instance.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// uses an input message to explain to the user what type of input is required.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_intro1.rs
    /// #
    /// use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write(1, 0, "Enter rating in cell D2:")?;
    ///
    ///     let data_validation = DataValidation::new()
    ///         .allow_whole_number(DataValidationRule::Between(1, 5))
    ///         .set_input_title("Enter a star rating!")?
    ///         .set_input_message("Enter rating 1-5.\nWhole numbers only.")?
    ///         .set_error_title("Value outside allowed range")?
    ///         .set_error_message("The input value must be an integer in the range 1-5.")?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    ///
    ///     // Save the file.
    ///     workbook.save("data_validation.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/data_validation_intro1.png">
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts users to a selection of values from a dropdown list.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_list_strings.rs
    /// #
    /// use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     worksheet.write(1, 0, "Select value in cell D2:")?;
    ///
    ///     let data_validation =
    ///         DataValidation::new().allow_list_strings(&["Pass", "Fail", "Incomplete"])?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    ///
    ///     // Save the file.
    ///     workbook.save("data_validation.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_strings.png">
    ///
    pub fn add_data_validation(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        data_validation: &DataValidation,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        let mut data_validation = data_validation.clone();

        // The "Any" validation type should be ignored if it doesn't have any
        // input or error titles or messages. This is the same rule as Excel.
        if data_validation.is_invalid_any() {
            return Ok(self);
        }

        // Store the data validation based on its range.
        let mut cell_range = utility::cell_range(first_row, first_col, last_row, last_col);
        if !data_validation.multi_range.is_empty() {
            cell_range.clone_from(&data_validation.multi_range);
        }

        self.data_validations.insert(cell_range, data_validation);

        Ok(self)
    }

    /// Add a sparkline to a worksheet cell.
    ///
    /// Sparklines are a feature of Excel 2010+ which allows you to add small
    /// charts to worksheet cells. These are useful for showing data trends in a
    /// compact visual format.
    ///
    /// The `add_sparkline()` method allows you to add a sparkline to a single
    /// cell that displays data from a 1D range of cells.
    ///
    /// The sparkline can be configured with all the parameters supported by
    /// Excel. See [`Sparkline`] for details.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `sparkline`: The [`Sparkline`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::SparklineError`] - An error that is raised when there is
    ///   an parameter error with the sparkline.
    /// - [`XlsxError::ChartError`] - An error that is raised when there is an
    ///   parameter error with the data range for the sparkline.
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name in chart range
    ///   cannot be blank.
    /// - [`XlsxError::SheetnameLengthExceeded`] - Worksheet name in chart range
    ///   exceeds Excel's limit of 31 characters.
    /// - [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name in
    ///   chart range cannot contain invalid characters: `[ ] : * ? / \`
    /// - [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name in
    ///   chart range cannot start or end with an apostrophe.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a sparkline to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_add_sparkline.rs
    /// #
    /// # use rust_xlsxwriter::{Sparkline, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add some sample data to plot.
    ///     worksheet.write_row(0, 0, [-2, 2, 3, -1, 0])?;
    ///
    ///     // Create a default line sparkline that plots the 1D data range.
    ///     let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 0, 4));
    ///
    ///     // Add it to the worksheet.
    ///     worksheet.add_sparkline(0, 5, &sparkline)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_add_sparkline.png">
    ///
    pub fn add_sparkline(
        &mut self,
        row: RowNum,
        col: ColNum,
        sparkline: &Sparkline,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check that the sparkline has a range.
        if !sparkline.data_range.has_data() {
            return Err(XlsxError::SparklineError(
                "Sparkline data range not set".to_string(),
            ));
        }

        // Check that the sparkline range is valid.
        sparkline.data_range.validate()?;

        // Check that the sparkline range is 1D.
        if !sparkline.data_range.is_1d() {
            let range = sparkline.data_range.error_range();
            return Err(XlsxError::SparklineError(format!(
                "Sparkline data range '{range}' must be a 1D range"
            )));
        }

        // Clone the sparkline and set a data range.
        let mut sparkline = sparkline.clone();
        sparkline.add_cell_range(row, col);

        // Store the sparkline.
        self.sparklines.push(sparkline);

        // Set some global worksheet flags.
        self.use_x14_extensions = true;
        self.has_sparklines = true;

        Ok(self)
    }

    /// Add a sparkline group to a worksheet range.
    ///
    /// Sparklines are a feature of Excel 2010+ which allows you to add small
    /// charts to worksheet cells. These are useful for showing data trends in a
    /// compact visual format.
    ///
    /// In Excel sparklines can be added as a single entity in a cell that
    /// refers to a 1D data range or as a "group" sparkline that is applied
    /// across a 1D range and refers to data in a 2D range. A grouped sparkline
    /// uses one sparkline for the specified range and any changes to it are
    /// applied to the entire sparkline group.
    ///
    /// The [`Worksheet::add_sparkline()`] method shown above allows you to add
    /// a sparkline to a single cell that displays data from a 1D range of cells
    /// whereas `add_sparkline_group()` applies the group sparkline to a range.
    ///
    /// The sparkline can be configured with all the parameters supported by
    /// Excel. See [`Sparkline`] for details.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `sparkline`: The [`Sparkline`] to insert into the cell.
    ///
    /// # Errors
    ///
    /// # Errors
    ///
    /// - [`XlsxError::SparklineError`] - An error that is raised when there is
    ///   an parameter error with the sparkline.
    /// - [`XlsxError::ChartError`] - An error that is raised when there is an
    ///   parameter error with the data range for the sparkline.
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name in chart range
    ///   cannot be blank.
    /// - [`XlsxError::SheetnameLengthExceeded`] - Worksheet name in chart range
    ///   exceeds Excel's limit of 31 characters.
    /// - [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name in
    ///   chart range cannot contain invalid characters: `[ ] : * ? / \`
    /// - [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name in
    ///   chart range cannot start or end with an apostrophe.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a sparkline group to a
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_add_sparkline_group.rs
    /// #
    /// # use rust_xlsxwriter::{Sparkline, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add some sample data to plot.
    ///     let data = [
    ///         [-2,  2,  3, -1,  0],
    ///         [30, 20, 33, 20, 15],
    ///         [1,  -1, -1,  1, -1]
    ///     ];
    ///     worksheet.write_row_matrix(0, 0, data)?;
    ///
    ///     // Create a default line sparkline that plots the 2D data range.
    ///     let sparkline = Sparkline::new().set_range(("Sheet1", 0, 0, 2, 4));
    ///
    ///     // Add it to the worksheet as a sparkline group.
    ///     worksheet.add_sparkline_group(0, 5, 2, 5, &sparkline)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_add_sparkline_group.png">
    ///
    pub fn add_sparkline_group(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        sparkline: &Sparkline,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Check that the sparkline has a range.
        if !sparkline.data_range.has_data() {
            return Err(XlsxError::SparklineError(
                "Sparkline data range not set".to_string(),
            ));
        }

        // Check that the sparkline range is valid.
        sparkline.data_range.validate()?;

        // Check that the sparkline range is 2D.
        if sparkline.data_range.is_1d() {
            let range = sparkline.data_range.error_range();
            return Err(XlsxError::SparklineError(format!(
                "Sparkline data range '{range}' must be a 2D range"
            )));
        }

        // Check that the group data range matches 1 dimension of the sparkline
        // data range.
        let row_range = (last_row - first_row + 1) as usize;
        let col_range = (last_col - first_col + 1) as usize;
        let num_cells = std::cmp::max(row_range, col_range);
        let (num_rows, num_cols) = sparkline.data_range.number_of_range_points();
        if num_cells != num_rows && num_cells != num_cols {
            let cell_range = format!("({first_row}, {first_col}, {last_row}, {last_col})");
            let sparkline_range = sparkline.data_range.error_range();
            return Err(XlsxError::SparklineError(format!(
                "Sparkline group range '{cell_range}' doesn't match dimensions of data range '{sparkline_range}'"
            )));
        }

        // Clone the sparkline and set a data range.
        let mut sparkline = sparkline.clone();
        sparkline.add_group_range(first_row, first_col, last_row, last_col);

        // Store the sparkline.
        self.sparklines.push(sparkline);

        // Set some global worksheet flags.
        self.use_x14_extensions = true;
        self.has_sparklines = true;

        Ok(self)
    }

    /// Protect a worksheet from modification.
    ///
    /// The `protect()` method protects a worksheet from modification. It works
    /// by enabling a cell's `locked` and `hidden` properties, if they have been
    /// set. A **locked** cell cannot be edited and this property is on by
    /// default for all cells. A **hidden** cell will display the results of a
    /// formula but not the formula itself.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/protection_alert.png">
    ///
    /// These properties can be set using the
    /// [`Format::set_locked()`](Format::set_locked)
    /// [`Format::set_unlocked()`](Format::set_unlocked) and
    /// [`Worksheet::set_hidden()`](Format::set_hidden) format methods. All cells
    /// have the `locked` property turned on by default (see the example below)
    /// so in general you don't have to explicitly turn it on.
    ///
    /// # Examples
    ///
    /// Example of cell locking and formula hiding in an Excel worksheet
    /// `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/app_worksheet_protection.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create some format objects.
    ///     let unlocked = Format::new().set_unlocked();
    ///     let hidden = Format::new().set_hidden();
    ///
    ///     // Protect the worksheet to turn on cell locking.
    ///     worksheet.protect();
    ///
    ///     // Examples of cell locking and hiding.
    ///     worksheet.write_string(0, 0, "Cell B1 is locked. It cannot be edited.")?;
    ///     worksheet.write_formula(0, 1, "=1+2")?; // Locked by default.
    ///
    ///     worksheet.write_string(1, 0, "Cell B2 is unlocked. It can be edited.")?;
    ///     worksheet.write_formula_with_format(1, 1, "=1+2", &unlocked)?;
    ///
    ///     worksheet.write_string(2, 0, "Cell B3 is hidden. The formula isn't visible.")?;
    ///     worksheet.write_formula_with_format(2, 1, "=1+2", &hidden)?;
    /// #
    /// #     worksheet.write_string(4, 0, "Use Menu -> Review -> Unprotect Sheet")?;
    /// #     worksheet.write_string(5, 0, "to remove the worksheet protection.")?;
    /// #
    /// #     worksheet.autofit();
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet_protection.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_worksheet_protection.png">
    ///
    pub fn protect(&mut self) -> &mut Worksheet {
        self.protection_on = true;

        self
    }

    /// Protect a worksheet from modification with a password.
    ///
    /// The `protect_with_password()` method is like the
    /// [`Worksheet::protect()`] method, see above, except that you can add an
    /// optional, weak, password to prevent modification.
    ///
    /// **Note**: Worksheet level passwords in Excel offer very weak protection.
    /// They do not encrypt your data and are very easy to deactivate. Full
    /// workbook encryption is not supported by `rust_xlsxwriter`. However, it
    /// is possible to encrypt an `rust_xlsxwriter` file using a third party
    /// open source tool called
    /// [msoffice-crypt](https://github.com/herumi/msoffice). This works for
    /// macOS, Linux and Windows:
    ///
    /// ```text
    /// msoffice-crypt.exe -e -p password clear.xlsx encrypted.xlsx
    /// ```
    ///
    /// # Parameters
    ///
    /// - `password`: The password string. Note, only ascii text passwords are
    ///   supported. Passing the empty string "" is the same as turning on
    ///   protection without a password.
    ///
    /// # Examples
    ///
    /// The following example demonstrates protecting a worksheet from editing
    /// with a password.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_protect_with_password.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Protect the worksheet from modification.
    ///     worksheet.protect_with_password("abc123");
    /// #
    /// #     worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_password.png">
    ///
    pub fn protect_with_password(&mut self, password: &str) -> &mut Worksheet {
        self.protection_on = true;
        self.protection_hash = utility::hash_password(password);

        self
    }

    /// Specify which worksheet elements should, or shouldn't, be protected.
    ///
    /// The `protect_with_password()` method is like the
    /// [`Worksheet::protect()`] method, see above, except it also specifies
    /// which worksheet elements should, or shouldn't, be protected.
    ///
    /// You can specify which worksheet elements protection should be on or off
    /// via a [`ProtectionOptions`] struct reference. The Excel options with
    /// their default states are shown below:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options1.png">
    ///
    /// # Parameters
    ///
    /// `options` - Worksheet protection options as defined by a
    /// [`ProtectionOptions`] struct reference.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet properties to
    /// be protected in a protected worksheet. In this case we protect the
    /// overall worksheet but allow columns and rows to be inserted.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_protect_with_options.rs
    /// #
    /// # use rust_xlsxwriter::{ProtectionOptions, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set some of the options and use the defaults for everything else.
    ///     let options = ProtectionOptions {
    ///         insert_columns: true,
    ///         insert_rows: true,
    ///         ..ProtectionOptions::default()
    ///     };
    ///
    ///     // Set the protection options.
    ///     worksheet.protect_with_options(&options);
    /// #
    /// #     worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Excel dialog for the output file, compare this with the default image
    /// above:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options2.png">
    ///
    pub fn protect_with_options(&mut self, options: &ProtectionOptions) -> &mut Worksheet {
        self.protection_on = true;
        self.protection_options = options.clone();

        self
    }

    /// Unprotect a range of cells in a protected worksheet.
    ///
    /// As shown in the example for the [`Worksheet::protect()`] method it is
    /// possible to unprotect a cell by setting the format `unprotect` property.
    /// Excel also offers an interface to unprotect larger ranges of cells. This
    /// is replicated in `rust_xlsxwriter` using the `unprotect_range()` method,
    /// see the example below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates unprotecting ranges in a protected
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_unprotect_range.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Protect the worksheet from modification.
    ///     worksheet.protect();
    ///
    ///     // Unprotect range D4:F10.
    ///     worksheet.unprotect_range(4, 3, 9, 5)?;
    ///
    ///     // Unprotect single cell B3 by repeating (row, col).
    ///     worksheet.unprotect_range(2, 1, 2, 1)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Dialog from the output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_unprotect_range.png">
    ///
    pub fn unprotect_range(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.unprotect_range_with_options(first_row, first_col, last_row, last_col, "", "")
    }

    /// Unprotect a range of cells in a protected worksheet, with options.
    ///
    /// This method is similar to [`Worksheet::unprotect_range()`], see above,
    /// expect that it allows you to specify two additional parameters to set
    /// the name of the range (instead of the default `Range1` .. `RangeN`) and
    /// also a optional weak password (see
    /// [`Worksheet::protect_with_password()`] for an explanation of what weak
    /// means here).
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `name`: The name of the range instead of `RangeN`. Can be blank if not
    ///   required.
    /// - `password`: The password to prevent modification of the range. Can be
    ///   blank if not required.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates unprotecting ranges in a protected
    /// worksheet, with additional options.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_unprotect_range_with_options.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Protect the worksheet from modification.
    ///     worksheet.protect();
    ///
    ///     // Unprotect range D4:F10 and give it a user defined name.
    ///     worksheet.unprotect_range_with_options(4, 3, 9, 5, "MyRange", "")?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Dialog from the output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_unprotect_range_with_options.png">
    ///
    pub fn unprotect_range_with_options(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        name: &str,
        password: &str,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        let range = utility::cell_range(first_row, first_col, last_row, last_col);
        let mut name = name.to_string();
        let password_hash = utility::hash_password(password);

        if name.is_empty() {
            name = format!("Range{}", 1 + self.unprotected_ranges.len());
        }

        self.unprotected_ranges.push((range, name, password_hash));

        Ok(self)
    }

    /// Set the selected cell or cells in a worksheet.
    ///
    /// The `set_selection()` method can be used to specify which cell or range
    /// of cells is selected in a worksheet. The most common requirement is to
    /// select a single cell, in which case the `first_` and `last_` parameters
    /// should be the same.
    ///
    /// The active cell within a selected range is determined by the order in
    /// which `first_` and `last_` are specified.
    ///
    /// Only one range of cells can be selected. The default cell selection is
    /// (0, 0, 0, 0), "A1".
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates selecting cells in worksheets. The order
    /// of selection within the range depends on the order of `first` and `last`.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_selection.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = workbook.add_worksheet();
    ///     worksheet1.set_selection(3, 2, 3, 2)?; // Cell C4
    ///
    ///     let worksheet2 = workbook.add_worksheet();
    ///     worksheet2.set_selection(3, 2, 6, 6)?; // Cells C4 to G7.
    ///
    ///     let worksheet3 = workbook.add_worksheet();
    ///     worksheet3.set_selection(6, 6, 3, 2)?; // Cells G7 to C4.
    ///
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_selection.png">
    pub fn set_selection(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // The first/last order can be reversed to allow a selection to go from
        // the end to the start. We take the active cell from the user first
        // row/col and then reverse them as required for the full range.
        let active_cell = utility::row_col_to_cell(first_row, first_col);

        let mut first_row = first_row;
        let mut first_col = first_col;
        let mut last_row = last_row;
        let mut last_col = last_col;

        if first_row > last_row {
            std::mem::swap(&mut first_row, &mut last_row);
        }

        if first_col > last_col {
            std::mem::swap(&mut first_col, &mut last_col);
        }

        let range = utility::cell_range(first_row, first_col, last_row, last_col);

        self.selected_range = (active_cell, range);

        Ok(self)
    }

    /// Set the first visible cell at the top left of a worksheet.
    ///
    /// This `set_top_left_cell()` method can be used to set the top leftmost
    /// visible cell in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the top and leftmost visible
    /// cell in the worksheet. Often used in conjunction with `set_selection()`
    /// to activate the same cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_top_left_cell.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set top-left cell to AA32.
    ///     worksheet.set_top_left_cell(31, 26)?;
    ///
    ///     // Also make this the active/selected cell.
    ///     worksheet.set_selection(31, 26, 31, 26)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_top_left_cell.png">
    ///
    pub fn set_top_left_cell(
        &mut self,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Ignore cell (0, 0) since that is the default top-left cell.
        if row == 0 && col == 0 {
            return Ok(self);
        }

        self.top_left_cell = utility::row_col_to_cell(row, col);

        Ok(self)
    }

    /// Write a user defined result to a worksheet formula cell.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using [`Worksheet::write_formula_with_format()`] or
    /// [`Worksheet::write_formula()`]. Instead it stores the value 0 as the
    /// formula result. It then sets a global flag in the xlsx file to say that
    /// all formulas and functions should be recalculated when the file is
    /// opened.
    ///
    /// This works fine with Excel and other spreadsheet applications. However,
    /// applications that don’t have a facility to calculate formulas will only
    /// display the 0 results.
    ///
    /// If required, it is possible to specify the calculated result of a
    /// formula using the `set_formula_result()` method.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `result`: The formula result to write to the cell.
    ///
    /// # Warnings
    ///
    /// You will get a warning if you try to set a formula result for a cell
    /// that doesn't have a formula.
    ///
    /// # Examples
    ///
    /// The following example demonstrates manually setting the result of a
    /// formula. Note, this is only required for non-Excel applications that
    /// don't calculate formula results.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_formula_result.rs
    /// #
    /// # use rust_xlsxwriter::{Formula, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Using string syntax.
    ///     worksheet
    ///         .write_formula(0, 0, "1+1")?
    ///         .set_formula_result(0, 0, "2");
    ///
    ///     // Or using a Formula type.
    ///     worksheet.write_formula(1, 0, Formula::new("2+2").set_result("4"))?;
    /// #
    /// #     workbook.save("formulas.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_formula_result(
        &mut self,
        row: RowNum,
        col: ColNum,
        result: impl Into<String>,
    ) -> &mut Worksheet {
        let lookup_row = if self.use_constant_memory {
            self.current_row
        } else {
            row
        };

        if let Some(columns) = self.data_table.get_mut(&lookup_row) {
            if let Some(cell) = columns.get_mut(&col) {
                match cell {
                    CellType::Formula {
                        result: cell_result,
                        ..
                    }
                    | CellType::ArrayFormula {
                        result: cell_result,
                        ..
                    } => {
                        *cell_result = Box::from(result.into());
                    }
                    _ => {
                        eprintln!("Cell ({row}, {col}) doesn't contain a formula.");
                    }
                }
            }
        }

        self
    }

    /// Write the default formula result for worksheet formulas.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using [`Worksheet::write_formula_with_format()`] or
    /// [`Worksheet::write_formula()`]. Instead it stores the value 0 as the
    /// formula result. It then sets a global flag in the xlsx file to say that
    /// all formulas and functions should be recalculated when the file is
    /// opened.
    ///
    /// However, for `LibreOffice` the default formula result should be set to
    /// the empty string literal `""`, via the `set_formula_result_default()`
    /// method, to force calculation of the result.
    ///
    /// # Parameters
    ///
    /// - `result`: The default formula result to write to the cell.
    ///
    /// # Examples
    ///
    /// The following example demonstrates manually setting the default result
    /// for all non-calculated formulas in a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_formula_result_default.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.set_formula_result_default("");
    /// #
    /// #     workbook.save("formulas.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_formula_result_default(&mut self, result: impl Into<String>) -> &mut Worksheet {
        self.default_result = Box::from(result.into());
        self
    }

    // -----------------------------------------------------------------------
    // Worksheet overlay/formatting methods.
    // -----------------------------------------------------------------------

    /// Add formatting to a cell without overwriting the cell data.
    ///
    /// In Excel the data in a worksheet cell is comprised of a type, a value
    /// and a format. When using `rust_xlsxwriter` the type is inferred and the
    /// value and format are generally written at the same time using methods
    /// like [`Worksheet::write_with_format()`]. However, if required you can
    /// write the data separately and then add the format using methods like
    /// `set_cell_format()`.
    ///
    /// Although this method requires an additional step it allows for use cases
    /// where it is easier to write a large amount of data in one go and then
    /// figure out where formatting should be applied. See also the
    /// documentation section on [Worksheet Cell
    /// formatting](../worksheet/index.html#cell-formatting).
    ///
    /// For use cases where the cell formatting changes based on cell values
    /// Conditional Formatting may be a better option (see [Working with
    /// Conditional Formats](../conditional_format/index.html)).
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format of a worksheet
    /// cell separately from writing the cell data.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_cell_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some formats.
    /// #     let red = Format::new().set_font_color("#FF0000");
    /// #     let green = Format::new().set_font_color("#00FF00");
    /// #
    ///     // Write an array of data.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///     worksheet.write_row_matrix(1, 1, data)?;
    ///
    ///     // Add formatting to some of the cells.
    ///     worksheet.set_cell_format(1, 1, &red)?;
    ///     worksheet.set_cell_format(2, 3, &green)?;
    ///     worksheet.set_cell_format(3, 5, &red)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_cell_format.png">
    ///
    pub fn set_cell_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let xf_index = self.format_xf_index(format);

        // Insert the format in a new or existing cell.
        self.insert_cell_format(row, col, xf_index);

        Ok(self)
    }

    /// Add formatting to a range of cells without overwriting the cell data.
    ///
    ///
    /// In Excel the data in a worksheet cell is comprised of a type, a value
    /// and a format. When using `rust_xlsxwriter` the type is inferred and the
    /// value and format are generally written at the same time using methods
    /// like [`Worksheet::write_with_format()`]. However, if required you can
    /// write the data separately and then add the format using methods like
    /// `set_range_format()` or [`Worksheet::set_cell_format()`] (see above).
    ///
    /// Although this method requires an additional step it allows for use cases
    /// where it is easier to write a large amount of data in one go and then
    /// figure out where formatting should be applied. See also the
    /// documentation section on [Worksheet Cell
    /// formatting](../worksheet/index.html#cell-formatting).
    ///
    /// For use cases where the cell formatting changes based on cell values
    /// Conditional Formatting may be a better option (see [Working with
    /// Conditional Formats](../conditional_format/index.html)).
    ///
    /// Note, this method should **not** be used to set the formatting for an
    /// entire worksheet since it would add an XML element for each of the 34
    /// billion cells in the worksheet which would result in a very large and
    /// slow output file. To set the format for the entire worksheet see the
    /// [`Worksheet::set_column_range_format()`] method.
    ///
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format of worksheet cells
    /// separately from writing the cell data.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_range_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add a format.
    ///     let border = Format::new().set_border(FormatBorder::Thin);
    ///
    ///     // Write an array of data.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///     worksheet.write_row_matrix(1, 1, data)?;
    ///
    ///     // Add formatting to the cells.
    ///     worksheet.set_range_format(1, 1, 3, 5, &border)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_range_format.png">
    ///
    pub fn set_range_format(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Get the index of the format object.
        let xf_index = self.format_xf_index(format);

        // Insert the format in a new or existing cells.
        for row in first_row..=last_row {
            for col in first_col..=last_col {
                self.insert_cell_format(row, col, xf_index);
            }
        }

        Ok(self)
    }

    /// Add formatting to a range of cells with an external border.
    ///
    /// This method is similar to the  [`Worksheet::set_range_format()`] method
    /// (see above) except it also adds a border around the cell range.
    ///
    /// Add a border around a range of cells in Excel is generally easy to do
    /// using the GUI interface. However, creating a border around a range of
    /// cells programmatically is much harder since it requires the creation of
    /// up to 9 separate formats and the tracking of where cells are relative to
    /// the border.
    ///
    /// The `set_range_format_with_border()` method is provided to simplify this
    /// task. It allows you to specify one format for the cells and another for
    /// the border.
    ///
    /// You should also consider formatting a range of cells as a Worksheet
    /// Table may be a better option than simple cell formatting (see the
    /// [`Table`] section of the documentation).
    ///
    /// Note, this method should **not** be used to set the formatting for an
    /// entire worksheet since it would add an XML element for each of the 34
    /// billion cells in the worksheet which would result in a very large and
    /// slow output file. To set the format for the entire worksheet see the
    /// [`Worksheet::set_column_range_format()`] method.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `cell_format`: The [`Format`] property for the cells in the range. If
    ///   you don't require internal formatting you can use `Format::default()`.
    /// - `border_format`: The [`Format`] property for the border. Only the
    ///   [`Format::set_border()`] and [`Format::set_border_color()`] properties
    ///   are used.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the format of worksheet cells
    /// separately from writing the cell data.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_range_format_with_border.rs
    /// #
    /// # use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add some formats.
    ///     let inner_border = Format::new().set_border(FormatBorder::Thin);
    ///     let outer_border = Format::new().set_border(FormatBorder::Double);
    ///
    ///     // Write an array of data.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///     worksheet.write_row_matrix(1, 1, data)?;
    ///
    ///     // Add formatting to the cells.
    ///     worksheet.set_range_format_with_border(1, 1, 3, 5, &inner_border, &outer_border)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_range_format_with_border.png">
    ///
    pub fn set_range_format_with_border(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        cell_format: &Format,
        border_format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        if first_row == last_row && first_col == last_col {
            self.set_range_border_cell(first_row, first_col, cell_format, border_format)?;
        } else if first_row == last_row {
            self.set_range_border_row(first_row, first_col, last_col, cell_format, border_format)?;
        } else if first_col == last_col {
            self.set_range_border_col(first_row, last_row, first_col, cell_format, border_format)?;
        } else {
            self.set_range_border_range(
                first_row,
                last_row,
                first_col,
                last_col,
                cell_format,
                border_format,
            )?;
        }

        Ok(self)
    }

    /// Clear the data and formatting from a worksheet cell.
    ///
    /// This method can be used to clear data and formatting previously written
    /// to a worksheet cell using one of the worksheet `write()` methods.
    ///
    /// This can occasionally be useful for scenarios where it is easier to add
    /// data in bulk but then remove certain elements.
    ///
    /// This method only clears data, it doesn't clear images or conditional
    /// formatting, or other non-data elements.
    ///
    /// Note, this method doesn't return a [`Result`] or errors. Instructions to
    /// clear non-existent cells are simply ignored.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Examples
    ///
    /// The following example demonstrates clearing some previously written cell
    /// data and formatting from a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_clear_cell.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add a format.
    /// #     let format = Format::new().set_font_color("#FF0000");
    /// #
    ///     // Some array data to write.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///
    ///     // Write the array data as a series of rows.
    ///     worksheet.write_row_with_format(0, 0, data[0], &format)?;
    ///     worksheet.write_row_with_format(1, 0, data[1], &format)?;
    ///     worksheet.write_row_with_format(2, 0, data[2], &format)?;
    ///
    ///     // Clear the first and last cell in the written data.
    ///     worksheet.clear_cell(0, 0);
    ///     worksheet.clear_cell(2, 4);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_clear_cell.png">
    ///
    pub fn clear_cell(&mut self, row: RowNum, col: ColNum) -> &mut Worksheet {
        // If cell is outside the allowed range it doesn't need to be cleared.
        if !self.check_dimensions_only(row, col) {
            return self;
        }

        self.clear_cell_internal(row, col);

        self
    }

    /// Clear the formatting from a worksheet cell.
    ///
    /// This method can be used to clear the formatting previously added to a
    /// worksheet cell using one of the worksheet `write_with_format()` methods.
    ///
    /// This can occasionally be useful for scenarios where it is easier to add
    /// formatted data in bulk but then remove the formatting from certain
    /// elements.
    ///
    /// See also the [`Worksheet::set_cell_format()`] method for a similar
    /// method to change the format of a cell.
    ///
    /// Note, this method doesn't return a [`Result`] or errors. Instructions to
    /// clear non-existent cells are simply ignored.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Examples
    ///
    /// The following example demonstrates clearing the formatting from some
    /// previously written cells in a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_clear_cell_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add a format.
    /// #     let format = Format::new().set_font_color("#FF0000");
    /// #
    ///     // Some array data to write.
    ///     let data = [
    ///         [10, 11, 12, 13, 14],
    ///         [20, 21, 22, 23, 24],
    ///         [30, 31, 32, 33, 34],
    ///     ];
    ///
    ///     // Write the array data as a series of rows.
    ///     worksheet.write_row_with_format(0, 0, data[0], &format)?;
    ///     worksheet.write_row_with_format(1, 0, data[1], &format)?;
    ///     worksheet.write_row_with_format(2, 0, data[2], &format)?;
    ///
    ///     // Clear the format from the first and last cells in the data.
    ///     worksheet.clear_cell_format(0, 0);
    ///     worksheet.clear_cell_format(2, 4);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_clear_cell_format.png">
    ///
    pub fn clear_cell_format(&mut self, row: RowNum, col: ColNum) -> &mut Worksheet {
        // If cell is outside the allowed range it doesn't need to be cleared.
        if !self.check_dimensions_only(row, col) {
            return self;
        }

        self.clear_cell_format_internal(row, col);

        self
    }

    // -----------------------------------------------------------------------
    // Worksheet serde methods.
    // -----------------------------------------------------------------------

    /// Write a Serde serializable struct to a worksheet.
    ///
    /// This method can be used to serialize [Serde](https://serde.rs) enabled
    /// data structures into cells in a worksheet.
    ///
    /// See [Working with Serde](crate::serializer#working-with-serde) for
    /// background details on how serialization works with `rust_xlsxwriter`.
    ///
    /// When serializing structs `rust_xlsxwriter` needs to know location where
    /// the serialization starts and also the type and field names of the struct
    /// being serialized. The field names are used as headers and the type name
    /// allows for several distinct structs to be serialized to the same
    /// worksheet.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/serialize_intro1.png">
    ///
    /// The worksheet methods that perform this function fall into two types:
    /// methods which use deserialization to find the fields from the *type* and
    /// methods that use serialization to find the fields from an *instance of
    /// the type*. The deserialization methods are easier to use but require
    /// that the struct derives the Serde [`Deserialize`] trait as well as the
    /// [`Serialize`] trait. The serialization methods work for anything else.
    ///
    /// There available methods are.
    ///
    /// - [`Worksheet::deserialize_headers()`]: The simplest most direct method.
    ///   It only requires the type of struct that you wish to serialize and
    ///   that it derives the [`Deserialize`] and [`Serialize`] traits. The
    ///   library uses this to infer the struct name and fields (via
    ///   deserialization).
    ///
    /// - [`Worksheet::deserialize_headers_with_format()`]: This is similar to
    ///   the previous method but it allows you to add a cell format for the
    ///   headers.
    ///
    /// - [`Worksheet::deserialize_headers_with_options()`]: Similar to the
    ///   previous methods but also allows configuration of the headers and
    ///   fields via [`SerializeFieldOptions`].
    ///
    /// - [`Worksheet::serialize_headers()`]: Similar to the
    ///   `deserialize_headers()` method but it requires a concrete instance of
    ///   the type of struct that you wish to serialize. The library uses this
    ///   to infer the struct name and fields (via serialization). This method
    ///   only requires that the struct derives [`Serialize`].
    ///
    /// Once the headers are set up an subsequent calls to `serialize()` will
    /// write the struct data in rows beneath the header.
    ///
    ///
    /// # Parameters
    ///
    /// - `data_structure`: A reference to a struct that implements the
    ///   [`serde::Serializer`] trait.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde derived
    /// data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize.rs
    /// #
    /// use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// use serde::{Deserialize, Serialize};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add a simple format for the headers.
    ///     let format = Format::new().set_bold();
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
    ///     #[serde(rename_all = "PascalCase")]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let item1 = Produce {
    ///         fruit: "Peach",
    ///         cost: 1.05,
    ///     };
    ///     let item2 = Produce {
    ///         fruit: "Plum",
    ///         cost: 0.15,
    ///     };
    ///     let item3 = Produce {
    ///         fruit: "Pear",
    ///         cost: 0.75,
    ///     };
    ///
    ///     // Set up the start location and headers of the data to be serialized.
    ///     worksheet.deserialize_headers_with_format::<Produce>(0, 0, &format)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
    ///
    ///     // Save the file.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn serialize<T>(&mut self, data_structure: &T) -> Result<&mut Worksheet, XlsxError>
    where
        T: Serialize,
    {
        self.serialize_data_structure(data_structure)?;

        Ok(self)
    }

    /// Write the location and headers for data serialization.
    ///
    /// The [`Worksheet::serialize()`] method, above, serializes Serde derived
    /// structs to worksheet cells. However, before you serialize the data you
    /// need to set the position in the worksheet where the headers will be
    /// written and where serialized data will be written.
    ///
    /// See [Setting serialization
    /// headers](crate::serializer#setting-serialization-headers) for more
    /// information.
    ///
    /// See also [`Worksheet::deserialize_headers()`] which only requires the
    /// serializable type and not an actual instance. That method requires that
    /// your struct also derives "Deserialize".
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data_structure`: A reference to a struct that implements the
    ///   [`serde::Serializer`] trait.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers1.rs
    /// #
    /// use rust_xlsxwriter::{Workbook, XlsxError};
    /// use serde::Serialize;
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Serialize)]
    ///     #[serde(rename_all = "PascalCase")]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let item1 = Produce {
    ///         fruit: "Peach",
    ///         cost: 1.05,
    ///     };
    ///     let item2 = Produce {
    ///         fruit: "Plum",
    ///         cost: 0.15,
    ///     };
    ///     let item3 = Produce {
    ///         fruit: "Pear",
    ///         cost: 0.75,
    ///     };
    ///
    ///     // Set up the start location and headers of the data to be serialized using
    ///     // any temporary or valid instance.
    ///     worksheet.serialize_headers(0, 0, &item1)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
    ///
    ///     // Save the file.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers1.png">
    ///
    /// You can serialize the data to any valid region of the worksheet:
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers2.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// # use serde::Serialize;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a serializable struct.
    /// #     #[derive(Serialize)]
    /// #     #[serde(rename_all = "PascalCase")]
    /// #     struct Produce {
    /// #         fruit: &'static str,
    /// #         cost: f64,
    /// #     }
    /// #
    /// #     // Create some data instances.
    /// #     let item1 = Produce {
    /// #         fruit: "Peach",
    /// #         cost: 1.05,
    /// #     };
    /// #     let item2 = Produce {
    /// #         fruit: "Plum",
    /// #         cost: 0.15,
    /// #     };
    /// #     let item3 = Produce {
    /// #         fruit: "Pear",
    /// #         cost: 0.75,
    /// #     };
    /// #
    /// #     // Set up the start location and headers of the data to be serialized using
    /// #     // any temporary or valid instance.
    ///     worksheet.serialize_headers(1, 2, &item1)?;
    /// #
    /// #   // Serialize the data.
    /// #   worksheet.serialize(&item1)?;
    /// #   worksheet.serialize(&item2)?;
    /// #   worksheet.serialize(&item3)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("serialize.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers2.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn serialize_headers<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data_structure: &T,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Serialize,
    {
        self.serialize_headers_with_format(row, col, data_structure, &Format::default())
    }

    /// Write the location and headers for data serialization, with formatting.
    ///
    /// The [`Worksheet::serialize()`] method, above, serializes Serde derived
    /// structs to worksheet cells. However, before you serialize the data you
    /// need to set the position in the worksheet where the headers will be
    /// written and where serialized data will be written. This method also
    /// allows you to set the format for the headers.
    ///
    /// See [Setting serialization
    /// headers](crate::serializer#setting-serialization-headers) for more
    /// information.
    ///
    /// See also [`Worksheet::deserialize_headers_with_format()`] which only
    /// requires the serializable type and not an actual instance. That method
    /// requires that your struct also derives "Deserialize".
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data_structure`: A reference to a struct that implements the
    ///   [`serde::Serializer`] trait.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde derived
    /// data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers4.rs
    /// #
    /// use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// use serde::Serialize;
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add a simple format for the headers.
    ///     let format = Format::new().set_bold();
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Serialize)]
    ///     #[serde(rename_all = "PascalCase")]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let item1 = Produce {
    ///         fruit: "Peach",
    ///         cost: 1.05,
    ///     };
    ///     let item2 = Produce {
    ///         fruit: "Plum",
    ///         cost: 0.15,
    ///     };
    ///     let item3 = Produce {
    ///         fruit: "Pear",
    ///         cost: 0.75,
    ///     };
    ///
    ///     // Set up the start location and headers of the data to be serialized using
    ///     // any temporary or valid instance.
    ///     worksheet.serialize_headers_with_format(0, 0, &item1, &format)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
    ///
    ///     // Save the file.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn serialize_headers_with_format<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data_structure: &T,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Serialize,
    {
        // Serialize the struct to determine the type name and the fields.
        let mut headers = SerializerHeader {
            struct_name: String::new(),
            field_names: vec![],
        };

        data_structure.serialize(&mut headers)?;

        self.store_serialization_headers(row, col, &headers, format)
    }

    /// Write the location and headers for data serialization, with additional
    /// options.
    ///
    /// The [`Worksheet::serialize()`] and
    /// [`Worksheet::serialize_headers_with_format()`] methods, above, set the
    /// serialization headers and location via an instance of the structure to
    /// be serialized. This will work for the majority of use cases, and for
    /// other cases you can adjust the output by using Serde Container or Field
    /// [Attributes]. See [Working with
    /// Serde](crate::serializer#working-with-serde).
    ///
    /// [Attributes]: https://serde.rs/attributes.html
    ///
    /// If these methods don't give you the output or flexibility you require
    /// you can use the `serialize_headers_with_options()` method with
    /// [`SerializeFieldOptions`] and [`CustomSerializeField`] options. This
    /// allows you to reorder, rename, format or skip headers and also define
    /// formatting for field values.
    ///
    /// See [`SerializeFieldOptions`] and [`CustomSerializeField`] for
    /// additional information and examples.
    ///
    /// See also [`Worksheet::deserialize_headers_with_options()`] which only
    /// requires the serializable type and not an actual instance. That method
    /// requires that your struct also derives "Deserialize".
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `data_structure`: A reference to a struct that implements the
    ///   [`serde::Serializer`] trait.
    /// - `header_options`: A [`SerializeFieldOptions`] instance.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_with_options.rs
    /// #
    /// use rust_xlsxwriter::{
    ///     CustomSerializeField, Format, SerializeFieldOptions, Workbook, XlsxError
    /// };
    /// use serde::Serialize;
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some formats to use with the serialization data.
    ///     let bold = Format::new().set_bold();
    ///     let currency = Format::new().set_num_format("$0.00");
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Serialize)]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let items = [
    ///         Produce {
    ///             fruit: "Peach",
    ///             cost: 1.05,
    ///         },
    ///         Produce {
    ///             fruit: "Plum",
    ///             cost: 0.15,
    ///         },
    ///         Produce {
    ///             fruit: "Pear",
    ///             cost: 0.75,
    ///         },
    ///     ];
    ///
    ///     // Set up the start location and headers of the data to be serialized using
    ///     // custom headers.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("fruit")
    ///             .rename("Fruit"),
    ///         CustomSerializeField::new("cost")
    ///             .rename("Price")
    ///             .set_value_format(currency),
    ///     ];
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_header_format(bold)
    ///         .set_custom_headers(&custom_headers);
    ///
    ///     worksheet.serialize_headers_with_options(0, 0, &items[0], &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&items)?;
    ///
    ///     // Save the file.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_with_options.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn serialize_headers_with_options<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        data_structure: &T,
        header_options: &SerializeFieldOptions,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Serialize,
    {
        // Serialize the struct to determine the type name and the fields.
        let mut headers = SerializerHeader {
            struct_name: String::new(),
            field_names: vec![],
        };

        data_structure.serialize(&mut headers)?;
        self.store_serialization_headers_with_options(row, col, &headers, header_options)
    }

    /// Write the location and headers for data serialization.
    ///
    /// The [`Worksheet::serialize()`] method, above, serializes Serde derived
    /// structs to worksheet cells. However, before you serialize the data you
    /// need to set the position in the worksheet where the headers will be
    /// written and where serialized data will be written.
    ///
    /// See [Setting serialization
    /// headers](crate::serializer#setting-serialization-headers) for more
    /// information.
    ///
    /// See also [`Worksheet::serialize_headers()`] which requires an instance
    /// of the serializable type but doesn't require that your struct also
    /// derives "Deserialize".
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_deserialize_headers1.rs
    /// #
    /// use rust_xlsxwriter::{Workbook, XlsxError};
    /// use serde::{Deserialize, Serialize};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
    ///     #[serde(rename_all = "PascalCase")]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let item1 = Produce {
    ///         fruit: "Peach",
    ///         cost: 1.05,
    ///     };
    ///     let item2 = Produce {
    ///         fruit: "Plum",
    ///         cost: 0.15,
    ///     };
    ///     let item3 = Produce {
    ///         fruit: "Pear",
    ///         cost: 0.75,
    ///     };
    ///
    ///     // Set up the start location and headers of the data to be serialized.
    ///     worksheet.deserialize_headers::<Produce>(0, 0)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
    ///
    ///     // Save the file.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers1.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn deserialize_headers<'de, T>(
        &mut self,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Deserialize<'de>,
    {
        self.deserialize_headers_with_format::<T>(row, col, &Format::default())
    }

    /// Write the location and headers for data serialization, with formatting.
    ///
    /// The [`Worksheet::serialize()`] method, above, serializes Serde derived
    /// structs to worksheet cells. However, before you serialize the data you
    /// need to set the position in the worksheet where the headers will be
    /// written and where serialized data will be written. This method also
    /// allows you to set the format for the headers.
    ///
    /// See [Setting serialization
    /// headers](crate::serializer#setting-serialization-headers) for more
    /// information.
    ///
    /// See also [`Worksheet::serialize_headers_with_format()`] which requires
    /// an instance of the serializable type but doesn't require that your
    /// struct also derives "Deserialize".
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `format`: The [`Format`] property for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add a simple format for the headers.
    /// #     let format = Format::new().set_bold();
    /// #
    /// #     // Create a serializable struct.
    /// #     #[derive(Deserialize, Serialize)]
    /// #     #[serde(rename_all = "PascalCase")]
    /// #     struct Produce {
    /// #         fruit: &'static str,
    /// #         cost: f64,
    /// #     }
    /// #
    /// #     // Create some data instances.
    /// #     let item1 = Produce {
    /// #         fruit: "Peach",
    /// #         cost: 1.05,
    /// #     };
    /// #     let item2 = Produce {
    /// #         fruit: "Plum",
    /// #         cost: 0.15,
    /// #     };
    /// #     let item3 = Produce {
    /// #         fruit: "Pear",
    /// #         cost: 0.75,
    /// #     };
    /// #
    /// #     // Set up the start location and headers of the data to be serialized.
    /// #     worksheet.deserialize_headers_with_format::<Produce>(0, 0, &format)?;
    /// #
    /// #     // Serialize the data.
    /// #     worksheet.serialize(&item1)?;
    /// #     worksheet.serialize(&item2)?;
    /// #     worksheet.serialize(&item3)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("serialize.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize.png">
    ///
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn deserialize_headers_with_format<'de, T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Deserialize<'de>,
    {
        // Deserialize the struct to determine the type name and the fields.
        let headers = deserialize_headers::<T>();

        self.store_serialization_headers(row, col, &headers, format)
    }

    /// Write the location and headers for data serialization, with additional
    /// options.
    ///
    /// The [`Worksheet::serialize()`] and
    /// [`Worksheet::deserialize_headers_with_format()`] methods, above, set the
    /// serialization headers and location via an instance of the structure to
    /// be serialized. This will work for the majority of use cases, and for
    /// other cases you can adjust the output by using Serde Container or Field
    /// [Attributes]. [Working with
    /// Serde](crate::serializer#working-with-serde).
    ///
    /// [Attributes]: https://serde.rs/attributes.html
    ///
    /// If these methods don't give you the output or flexibility you require
    /// you can use the `deserialize_headers_with_options()` method with
    /// [`SerializeFieldOptions`] and [`CustomSerializeField`] options. This
    /// allows you to reorder, rename, format or skip headers and also define
    /// formatting for field values.
    ///
    /// See [`SerializeFieldOptions`] and [`CustomSerializeField`] for
    /// additional information and examples.
    ///
    /// See also [`Worksheet::serialize_headers_with_options()`] which requires
    /// an instance of the serializable type but doesn't require that your
    /// struct also derives "Deserialize".
    ///
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `header_options`: A [`SerializeFieldOptions`] instance.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_with_options2.rs
    /// #
    /// use rust_xlsxwriter::{
    ///     CustomSerializeField, Format, SerializeFieldOptions, Workbook, XlsxError
    /// };
    /// use serde::{Deserialize, Serialize};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Add some formats to use with the serialization data.
    ///     let bold = Format::new().set_bold();
    ///     let currency = Format::new().set_num_format("$0.00");
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let items = [
    ///         Produce {
    ///             fruit: "Peach",
    ///             cost: 1.05,
    ///         },
    ///         Produce {
    ///             fruit: "Plum",
    ///             cost: 0.15,
    ///         },
    ///         Produce {
    ///             fruit: "Pear",
    ///             cost: 0.75,
    ///         },
    ///     ];
    ///
    ///     // Set up the start location and headers of the data to be serialized using
    ///     // custom headers.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("fruit")
    ///             .rename("Fruit"),
    ///         CustomSerializeField::new("cost")
    ///             .rename("Price")
    ///             .set_value_format(currency),
    ///     ];
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_header_format(bold)
    ///         .set_custom_headers(&custom_headers);
    ///
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&items)?;
    ///
    ///     // Save the file.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_with_options.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn deserialize_headers_with_options<'de, T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        header_options: &SerializeFieldOptions,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: Deserialize<'de>,
    {
        // Deserialize the struct to determine the type name and the fields.
        let headers = deserialize_headers::<T>();

        self.store_serialization_headers_with_options(row, col, &headers, header_options)
    }

    /// Write the location and headers for data serialization.
    ///
    /// The [`Worksheet::serialize()`] method, above, serializes Serde derived
    /// structs to worksheet cells. However, before you serialize the data you
    /// need to set the position in the worksheet where the headers will be
    /// written and where serialized data will be written.
    ///
    /// See [Setting serialization
    /// headers](crate::serializer#setting-serialization-headers) for more
    /// information.
    ///
    /// See also [`Worksheet::serialize_headers()`] which requires an instance
    /// of the serializable type but doesn't require that your struct also
    /// derives `Deserialize`, and [`Worksheet::deserialize_headers()`] which
    /// does.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    /// - [`XlsxError::SerdeError`] - Errors encountered during the Serde
    ///   serialization.
    ///
    /// # Examples
    ///
    /// ```
    /// # // This code is available in examples/doc_xlsxserialize_intro.rs
    /// #
    /// use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
    /// use serde::Serialize;
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Create a serializable struct.
    ///     #[derive(XlsxSerialize, Serialize)]
    ///     #[xlsx(header_format = Format::new().set_bold())]
    ///     struct Produce {
    ///         #[xlsx(rename = "Item")]
    ///         #[xlsx(column_width = 12.0)]
    ///         fruit: &'static str,
    ///
    ///         #[xlsx(rename = "Price", num_format = "$0.00")]
    ///         cost: f64,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let item1 = Produce {
    ///         fruit: "Peach",
    ///         cost: 1.05,
    ///     };
    ///
    ///     let item2 = Produce {
    ///         fruit: "Plum",
    ///         cost: 0.15,
    ///     };
    ///
    ///     let item3 = Produce {
    ///         fruit: "Pear",
    ///         cost: 0.75,
    ///     };
    ///
    ///     // Set the serialization location and headers.
    ///     worksheet.set_serialize_headers::<Produce>(0, 0)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("serialize.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_intro.png">
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_serialize_headers<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError>
    where
        T: XlsxSerialize + Serialize,
    {
        let header_options = T::to_serialize_field_options();
        self.store_custom_serialization_headers(row, col, &header_options)
    }

    /// Get the row/column dimensions of a serialized area.
    ///
    /// When serializing data it generally isn't necessary to track the row and
    /// column range of the cells that are written since `rust_xlsxwriter` does
    /// it automatically. However, it is sometimes useful to know the range of
    /// the serialization after the data is written in order to refer to it in
    /// another function such as a conditional format or a chart.
    ///
    /// The `get_serialize_dimensions()` function returns the row/column
    /// dimensions of a serialized area for use cases where you need to know the
    /// range of the data that was written. The dimensions are returned as a
    /// `(min_row, min_col, max_row, max_col)` tuple in a `Result<>`.
    ///
    /// # Parameters
    ///
    /// - `struct_name`: The name/type of the target struct as a string.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::ParameterError`] - Unknown or unserialized struct name.
    ///
    /// # Examples
    ///
    /// Example of getting the dimensions of some serialized data. In this
    /// example we use the dimensions to set a conditional format range.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_dimensions1.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
    /// # };
    /// # use serde::Serialize;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a serializable struct.
    ///     #[derive(Serialize)]
    ///     #[serde(rename_all = "PascalCase")]
    ///     struct MyStruct {
    ///         col1: u8,
    ///         col2: u8,
    ///         col3: u8,
    ///         col4: u8,
    ///     }
    ///
    ///     // Create some sample data.
    /// #     #[rustfmt::skip]
    ///     let data = [
    ///         MyStruct {col1: 34,  col2: 73, col3: 39, col4: 32},
    ///         MyStruct {col1: 5,   col2: 24, col3: 1,  col4: 84},
    ///         MyStruct {col1: 28,  col2: 79, col3: 97, col4: 13},
    ///         MyStruct {col1: 27,  col2: 71, col3: 40, col4: 17},
    ///         MyStruct {col1: 88,  col2: 25, col3: 33, col4: 23},
    ///         MyStruct {col1: 23,  col2: 99, col3: 20, col4: 88},
    ///         MyStruct {col1: 7,   col2: 57, col3: 88, col4: 28},
    ///         MyStruct {col1: 53,  col2: 78, col3: 1,  col4: 96},
    ///         MyStruct {col1: 60,  col2: 54, col3: 81, col4: 66},
    ///         MyStruct {col1: 70,  col2: 5,  col3: 46, col4: 14},
    ///     ];
    ///
    ///     // Set the serialization location and headers.
    ///     worksheet.serialize_headers(0, 0, &data[1])?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&data)?;
    ///
    ///     // Add a format. Green fill with dark green text.
    ///     let format = Format::new()
    ///         .set_font_color("006100")
    ///         .set_background_color("C6EFCE");
    ///
    ///     // Create a conditional format.
    ///     let conditional_format = ConditionalFormatCell::new()
    ///         .set_rule(ConditionalFormatCellRule::GreaterThanOrEqualTo(50))
    ///         .set_format(format);
    ///
    ///     // Get the range that the serialization applies to.
    ///     let (min_row, min_col, max_row, max_col) =
    ///         worksheet.get_serialize_dimensions("MyStruct")?;
    ///
    ///     // Write the conditional format to the serialization area. Note, we add 1 to
    ///     // the minimum row number to skip the headers.
    ///     worksheet.add_conditional_format(
    ///         min_row + 1,
    ///         min_col,
    ///         max_row,
    ///         max_col,
    ///         &conditional_format,
    ///     )?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("serialize.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_dimensions1.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn get_serialize_dimensions(
        &mut self,
        struct_name: &str,
    ) -> Result<(RowNum, ColNum, RowNum, ColNum), XlsxError> {
        self.serializer_state.get_dimensions(struct_name)
    }

    /// Get the row/column dimensions of a column in a serialized area.
    ///
    /// When serializing data it generally isn't necessary to track the row and
    /// column range of the cells that are written since `rust_xlsxwriter` does
    /// it automatically. However, it is sometimes useful to know the range of
    /// the serialization after the data is written in order to refer to it in
    /// another function such as a conditional format or a chart.
    ///
    /// The `get_serialize_column_dimensions()` function returns the row/column
    /// dimensions of a field in a serialized area for use cases where you need to
    /// know the range of the data that was written. The dimensions are returned
    /// as a `(min_row, col, max_row, col)` tuple in a `Result<>`.
    ///
    /// # Parameters
    ///
    /// - `struct_name`: The name/type of the target struct, as a string.
    /// - `struct_name`: The name of the field in the target struct, as a
    ///   string.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::ParameterError`] - Unknown or unserialized struct name or
    ///   field.
    ///
    ///
    /// # Examples
    ///
    /// Example of getting the field/column dimensions of some serialized data. In
    /// this example we use the dimensions to set a conditional format range.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_dimensions2.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     ConditionalFormatCell, ConditionalFormatCellRule, Format, Workbook, XlsxError,
    /// # };
    /// # use serde::Serialize;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a serializable struct.
    ///     #[derive(Serialize)]
    ///     #[serde(rename_all = "PascalCase")]
    ///     struct MyStruct {
    ///         col1: u8,
    ///         col2: u8,
    ///         col3: u8,
    ///         col4: u8,
    ///     }
    ///
    ///     // Create some sample data.
    /// #     #[rustfmt::skip]
    ///     let data = [
    ///         MyStruct {col1: 34,  col2: 73, col3: 39, col4: 32},
    ///         MyStruct {col1: 5,   col2: 24, col3: 1,  col4: 84},
    ///         MyStruct {col1: 28,  col2: 79, col3: 97, col4: 13},
    ///         MyStruct {col1: 27,  col2: 71, col3: 40, col4: 17},
    ///         MyStruct {col1: 88,  col2: 25, col3: 33, col4: 23},
    ///         MyStruct {col1: 23,  col2: 99, col3: 20, col4: 88},
    ///         MyStruct {col1: 7,   col2: 57, col3: 88, col4: 28},
    ///         MyStruct {col1: 53,  col2: 78, col3: 1,  col4: 96},
    ///         MyStruct {col1: 60,  col2: 54, col3: 81, col4: 66},
    ///         MyStruct {col1: 70,  col2: 5,  col3: 46, col4: 14},
    ///     ];
    ///
    ///     // Set the serialization location and headers.
    ///     worksheet.serialize_headers(0, 0, &data[1])?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&data)?;
    ///
    ///     // Add a format. Green fill with dark green text.
    ///     let format = Format::new()
    ///         .set_font_color("006100")
    ///         .set_background_color("C6EFCE");
    ///
    ///     // Create a conditional format.
    ///     let conditional_format = ConditionalFormatCell::new()
    ///         .set_rule(ConditionalFormatCellRule::GreaterThanOrEqualTo(50))
    ///         .set_format(format);
    ///
    ///     // Get the range that the serialization field applies to. Note that we must
    ///     // match the Serde field name which has been renamed in PascalCase to Col3
    ///     // (not col3). Also note that min_col and max_col are the same in this case
    ///     // but we give them separate names for the sake of the example.
    ///     let (min_row, min_col, max_row, max_col) =
    ///         worksheet.get_serialize_column_dimensions("MyStruct", "Col3")?;
    ///
    ///     // Write the conditional format to the serialization area. Note, we add 1 to
    ///     // the minimum row number to skip the headers.
    ///     worksheet.add_conditional_format(
    ///         min_row + 1,
    ///         min_col,
    ///         max_row,
    ///         max_col,
    ///         &conditional_format,
    ///     )?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("serialize.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_serialize_dimensions2.png">
    ///
    #[cfg(feature = "serde")]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn get_serialize_column_dimensions(
        &mut self,
        struct_name: &str,
        field_name: &str,
    ) -> Result<(RowNum, ColNum, RowNum, ColNum), XlsxError> {
        self.serializer_state
            .get_column_dimensions(struct_name, field_name)
    }

    // Store serialization headers and options.
    #[cfg(feature = "serde")]
    fn store_serialization_headers_with_options(
        &mut self,
        row: RowNum,
        col: ColNum,
        headers: &SerializerHeader,
        header_options: &SerializeFieldOptions,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check that any custom field names match the actual field names.
        let field_names: HashSet<String> = HashSet::from_iter(headers.field_names.clone());
        for custom_header in &header_options.custom_headers {
            if !field_names.contains(&custom_header.field_name) {
                return Err(XlsxError::ParameterError(format!(
                    "No custom field name '{}' found for struct '{}'",
                    custom_header.field_name, headers.struct_name
                )));
            }
        }

        // Create a map of the user defined custom field settings to overwrite
        // the default field settings.
        let mut custom_fields: HashMap<&String, &CustomSerializeField> = HashMap::new();
        for custom_header in &header_options.custom_headers {
            custom_fields.insert(&custom_header.field_name, custom_header);
        }

        // Clone the header options to modify it and store it internally.
        let mut header_options = header_options.clone();
        header_options.struct_name.clone_from(&headers.struct_name);

        // Create a "custom" header for default fields or replace them with user
        // specified custom fields. The "use_custom_headers_only" overrides the
        // default headers to allow users to skip fields.
        if !header_options.use_custom_headers_only {
            let mut custom_headers: Vec<CustomSerializeField> = vec![];

            for field_name in &headers.field_names {
                match custom_fields.get(field_name) {
                    Some(custom_field) => {
                        if !custom_field.skip {
                            custom_headers.push((*custom_field).clone());
                        }
                    }
                    None => custom_headers.push(CustomSerializeField::new(field_name)),
                }
            }

            header_options.custom_headers = custom_headers;
        }

        self.store_custom_serialization_headers(row, col, &header_options)
    }

    // Store serialization headers with default options.
    #[cfg(feature = "serde")]
    fn store_serialization_headers(
        &mut self,
        row: RowNum,
        col: ColNum,
        headers: &SerializerHeader,
        header_format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Convert the field names to custom header structs.
        let custom_headers: Vec<CustomSerializeField> = headers
            .field_names
            .iter()
            .map(CustomSerializeField::new)
            .collect();

        // Transfer the options to a default option struct.
        let header_options = SerializeFieldOptions {
            struct_name: headers.struct_name.clone(),
            header_format: Some(header_format.clone()),
            custom_headers,
            ..Default::default()
        };

        self.store_custom_serialization_headers(row, col, &header_options)
    }

    // Write serialization headers to the worksheet.
    #[cfg(feature = "serde")]
    fn store_custom_serialization_headers(
        &mut self,
        row: RowNum,
        col: ColNum,
        header_options: &SerializeFieldOptions,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.

        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check for empty struct name.
        if header_options.struct_name.is_empty() {
            return Err(XlsxError::ParameterError(
                "Struct not found or serialized/deserialized.".to_string(),
            ));
        }

        // Check for empty struct members.
        if header_options.custom_headers.is_empty() {
            return Err(XlsxError::ParameterError(format!(
                "No members found/specified for struct '{}'",
                header_options.struct_name
            )));
        }

        let mut fields = HashMap::new();
        let min_row = row;
        let min_col = col;
        let mut max_row = row;
        let mut max_col = col;

        let col_initial = col;
        let write_headers = header_options.has_headers;

        let mut col_offset = 0;
        for custom_header in &header_options.custom_headers {
            if custom_header.skip {
                continue;
            }

            let col = col_initial + col_offset as u16;
            let mut custom_header = custom_header.clone();
            custom_header.col = col;
            max_col = col;
            col_offset += 1;

            // Set the column width if specified by user.
            if let Some(width) = custom_header.width {
                self.set_column_width(col, width)?;
            } else if let Some(pixel_width) = custom_header.pixel_width {
                self.set_column_width_pixels(col, pixel_width)?;
            }

            // Set the column format if specified by user.
            if let Some(format) = &custom_header.column_format {
                self.set_column_format(col, format)?;
            }

            // Use the column specific header format or else the header row
            // format, and if neither of those have been specified then write
            // without a format.
            if write_headers {
                if let Some(format) = &custom_header.header_format {
                    self.write_string_with_format(
                        max_row,
                        col,
                        &custom_header.header_name,
                        format,
                    )?;
                } else if let Some(format) = &header_options.header_format {
                    self.write_string_with_format(
                        max_row,
                        col,
                        &custom_header.header_name,
                        format,
                    )?;
                } else {
                    self.write_string(max_row, col, &custom_header.header_name)?;
                };
            }

            fields.insert(custom_header.field_name.clone(), custom_header);
        }

        // Start the data serialization one row down if headers were written.
        if write_headers {
            max_row += 1;
        }

        // If a previous serialization was carried out with the same struct name
        // then write the previous table formatting.
        if let Some(header_config) = self
            .serializer_state
            .structs
            .get_mut(&header_options.struct_name)
        {
            if let Some(table_data) = header_config.get_table() {
                self.write_serialized_table(&table_data)?;
            }
        }

        // Clone the new user defined table format, if present.
        let table = match &header_options.table {
            Some(table) => {
                let mut table = table.clone();
                if !header_options.has_headers {
                    table.show_header_row = false;
                }
                table.is_serde_table = true;
                Some(table)
            }
            None => None,
        };

        // Store meta data for the struct/headers.
        self.serializer_state.structs.insert(
            header_options.struct_name.clone(),
            SerializationHeaderConfig {
                fields,
                min_row,
                min_col,
                max_row,
                max_col,
                table,
            },
        );

        Ok(self)
    }

    // Serialize the parent data structure to the worksheet.
    #[cfg(feature = "serde")]
    fn serialize_data_structure<T>(&mut self, data_structure: &T) -> Result<(), XlsxError>
    where
        T: Serialize,
    {
        data_structure.serialize(self)?;
        Ok(())
    }

    // Serialize individual data items to a worksheet cell.
    #[cfg(feature = "serde")]
    pub(crate) fn serialize_to_worksheet_cell(
        &mut self,
        data: impl IntoExcelData,
    ) -> Result<(), XlsxError> {
        let result = self.serializer_state.current_state();

        match result {
            Ok(result) => {
                let (row, col, value_format) = result;
                match &*value_format {
                    Some(format) => self.write_with_format(row, col, data, format).map(|_| ()),
                    None => self.write(row, col, data).map(|_| ()),
                }
            }
            Err(()) => Ok(()),
        }
    }

    // Add any tables that were added as part of serialization formatting.
    #[cfg(feature = "serde")]
    pub(crate) fn store_serialized_tables(&mut self) -> Result<&mut Worksheet, XlsxError> {
        let tables = self.serializer_state.get_tables();

        for table_data in tables {
            self.write_serialized_table(&table_data)?;
        }

        Ok(self)
    }

    // Write a table that is part of serialization formatting.
    #[cfg(feature = "serde")]
    pub(crate) fn write_serialized_table(
        &mut self,
        table_data: &TableData,
    ) -> Result<&mut Worksheet, XlsxError> {
        let min_col = table_data.1;
        let max_col = table_data.3;
        let min_row = table_data.0;
        let mut max_row = table_data.2;
        let table = &table_data.4;

        if table.show_total_row {
            max_row += 1;
        }

        self.add_table(min_row, min_col, max_row, max_col, table)
    }

    // -----------------------------------------------------------------------
    // Worksheet page setup methods.
    // -----------------------------------------------------------------------

    /// Display the worksheet cells from right to left for some versions of
    /// Excel.
    ///
    /// The `set_right_to_left()` method is used to change the default direction
    /// of the worksheet from left-to-right, with the A1 cell in the top left,
    /// to right-to-left, with the A1 cell in the top right.
    ///
    /// This is useful when creating Arabic, Hebrew or other near or far eastern
    /// worksheets that use right-to-left as the default direction.
    ///
    /// Depending on your use case, and text, you may also need to use the
    /// [`Format::set_reading_direction()`](crate::Format::set_reading_direction)
    /// method to set the direction of the text within the cells.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates changing the default worksheet and
    /// cell text direction changed from left-to-right to right-to-left, as
    /// required by some middle eastern versions of Excel.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_right_to_left.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     // Add the cell formats.
    ///     let format_left_to_right = Format::new().set_reading_direction(1);
    ///     let format_right_to_left = Format::new().set_reading_direction(2);
    ///
    ///     // Add a worksheet in the standard left to right direction.
    ///     let worksheet1 = workbook.add_worksheet();
    ///
    ///     // Make the column wider for clarity.
    ///     worksheet1.set_column_width(0,25)?;
    ///
    ///     // Standard direction:         | A1 | B1 | C1 | ...
    ///     worksheet1.write(0, 0, "نص عربي / English text")?;
    ///     worksheet1.write_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    ///     worksheet1.write_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;
    ///
    ///     // Add a worksheet and change it to right to left direction.
    ///     let worksheet2 = workbook.add_worksheet();
    ///     worksheet2.set_right_to_left(true);
    ///
    ///     // Make the column wider for clarity.
    ///     worksheet2.set_column_width(0, 25)?;
    ///
    ///     // Right to left direction:    ... | C1 | B1 | A1 |
    ///     worksheet2.write(0, 0, "نص عربي / English text")?;
    ///     worksheet2.write_with_format(1, 0, "نص عربي / English text", &format_left_to_right)?;
    ///     worksheet2.write_with_format(2, 0, "نص عربي / English text", &format_right_to_left)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_right_to_left.png">
    ///
    pub fn set_right_to_left(&mut self, enable: bool) -> &mut Worksheet {
        self.right_to_left = enable;
        self
    }

    /// Make a worksheet the active/initially visible worksheet in a workbook.
    ///
    /// The `set_active()` method is used to specify which worksheet is
    /// initially visible in a multi-sheet workbook. If no worksheet is set then
    /// the first worksheet is made the active worksheet, like in Excel.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting a worksheet as the visible
    /// worksheet when a file is opened.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_active.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = Worksheet::new();
    ///     let worksheet3 = Worksheet::new();
    ///     let mut worksheet2 = Worksheet::new();
    ///
    ///     worksheet2.set_active(true);
    /// #
    /// #   workbook.push_worksheet(worksheet1);
    /// #   workbook.push_worksheet(worksheet2);
    /// #   workbook.push_worksheet(worksheet3);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_active.png">
    ///
    pub fn set_active(&mut self, enable: bool) -> &mut Worksheet {
        self.active = enable;

        // Activated worksheets must also be selected and cannot be hidden.
        if self.active {
            self.selected = true;
            self.visible = Visible::Default;
        }

        self
    }

    /// Set a worksheet tab as selected.
    ///
    /// The `set_selected()` method is used to indicate that a worksheet is
    /// selected in a multi-sheet workbook.
    ///
    /// A selected worksheet has its tab highlighted. Selecting worksheets is a
    /// way of grouping them together so that, for example, several worksheets
    /// could be printed in one go. A worksheet that has been activated via the
    /// [`Worksheet::set_active()`] method will also appear as selected.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates selecting worksheet in a workbook.
    /// The active worksheet is selected by default so in this example the first
    /// two worksheets are selected.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_selected.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = Worksheet::new();
    ///     let worksheet3 = Worksheet::new();
    ///     let mut worksheet2 = Worksheet::new();
    ///
    ///     worksheet2.set_selected(true);
    /// #
    /// #   workbook.push_worksheet(worksheet1);
    /// #   workbook.push_worksheet(worksheet2);
    /// #   workbook.push_worksheet(worksheet3);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_selected.png">
    ///
    pub fn set_selected(&mut self, enable: bool) -> &mut Worksheet {
        self.selected = enable;

        // Selected worksheets cannot be hidden.
        if self.selected {
            self.visible = Visible::Default;
        }

        self
    }

    /// Hide a worksheet.
    ///
    /// The `set_hidden()` method is used to hide a worksheet. This can be used
    /// to hide a worksheet in order to avoid confusing a user with intermediate
    /// data or calculations.
    ///
    /// In Excel a hidden worksheet can not be activated or selected so this
    /// method is mutually exclusive with the [`Worksheet::set_active()`] and
    /// [`Worksheet::set_selected()`] methods. In addition, since the first
    /// worksheet will default to being the active worksheet, you cannot hide
    /// the first worksheet without activating another sheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates hiding a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_hidden.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = Worksheet::new();
    ///     let worksheet3 = Worksheet::new();
    ///     let mut worksheet2 = Worksheet::new();
    ///
    ///     worksheet2.set_hidden(true);
    /// #
    /// #    workbook.push_worksheet(worksheet1);
    /// #    workbook.push_worksheet(worksheet2);
    /// #    workbook.push_worksheet(worksheet3);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_hidden.png">
    ///
    pub fn set_hidden(&mut self, enable: bool) -> &mut Worksheet {
        if enable {
            self.visible = Visible::Hidden;
        } else {
            self.visible = Visible::Default;
        }

        // Hidden worksheets cannot be active or hidden.
        if self.visible == Visible::Hidden {
            self.selected = false;
            self.active = false;
        }

        self
    }

    /// Hide a worksheet. Can only be unhidden in Excel by VBA.
    ///
    /// The `set_very_hidden()` method can be used to hide a worksheet similar
    /// to the [`Worksheet::set_hidden()`] method. The difference is that the
    /// worksheet cannot be unhidden in the the Excel user interface. The Excel
    /// worksheet `xlSheetVeryHidden` option can only be unset programmatically
    /// by VBA.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_very_hidden(&mut self, enable: bool) -> &mut Worksheet {
        if enable {
            self.visible = Visible::VeryHidden;
        } else {
            self.visible = Visible::Default;
        }

        // Hidden worksheets cannot be active or hidden.
        if self.visible == Visible::VeryHidden {
            self.selected = false;
            self.active = false;
        }

        self
    }

    /// Set current worksheet as the first visible sheet tab.
    ///
    /// The [`Worksheet::set_active()`]  method determines which worksheet is
    /// initially selected. However, if there are a large number of worksheets
    /// the selected worksheet may not appear on the screen. To avoid this you
    /// can select which is the leftmost visible worksheet tab using
    /// `set_first_tab()`.
    ///
    /// This method is not required very often. The default is the first
    /// worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_first_tab(&mut self, enable: bool) -> &mut Worksheet {
        self.first_sheet = enable;

        // First visible worksheet cannot be hidden.
        if self.selected {
            self.visible = Visible::Default;
        }
        self
    }

    /// Set the color of the worksheet tab.
    ///
    /// The `set_tab_color()` method can be used to change the color of the
    /// worksheet tab. This is useful for highlighting the important tab in a
    /// group of worksheets.
    ///
    /// # Parameters
    ///
    /// - `color`: The tab color property defined by a [`Color`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates set the tab color of worksheets.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_tab_color.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let mut worksheet1 = Worksheet::new();
    ///     let mut worksheet2 = Worksheet::new();
    ///     let mut worksheet3 = Worksheet::new();
    ///     let mut worksheet4 = Worksheet::new();
    ///
    ///     worksheet1.set_tab_color(Color::Red);
    ///     worksheet2.set_tab_color(Color::Green);
    ///     worksheet3.set_tab_color(Color::RGB(0xFF9900));
    ///
    ///     // worksheet4 will have the default color.
    ///     worksheet4.set_active(true);
    /// #
    /// #    workbook.push_worksheet(worksheet1);
    /// #    workbook.push_worksheet(worksheet2);
    /// #    workbook.push_worksheet(worksheet3);
    /// #    workbook.push_worksheet(worksheet4);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_tab_color.png">
    ///
    pub fn set_tab_color(&mut self, color: impl Into<Color>) -> &mut Worksheet {
        let color = color.into();
        if color.is_valid() {
            self.tab_color = color;
        }

        self
    }

    /// Set the paper type/size when printing.
    ///
    /// This method is used to set the paper format for the printed output of a
    /// worksheet. The following paper styles are available:
    ///
    /// | Index    | Paper format            | Paper size           |
    /// | :------- | :---------------------- | :------------------- |
    /// | 0        | Printer default         | Printer default      |
    /// | 1        | Letter                  | 8 1/2 x 11 in        |
    /// | 2        | Letter Small            | 8 1/2 x 11 in        |
    /// | 3        | Tabloid                 | 11 x 17 in           |
    /// | 4        | Ledger                  | 17 x 11 in           |
    /// | 5        | Legal                   | 8 1/2 x 14 in        |
    /// | 6        | Statement               | 5 1/2 x 8 1/2 in     |
    /// | 7        | Executive               | 7 1/4 x 10 1/2 in    |
    /// | 8        | A3                      | 297 x 420 mm         |
    /// | 9        | A4                      | 210 x 297 mm         |
    /// | 10       | A4 Small                | 210 x 297 mm         |
    /// | 11       | A5                      | 148 x 210 mm         |
    /// | 12       | B4                      | 250 x 354 mm         |
    /// | 13       | B5                      | 182 x 257 mm         |
    /// | 14       | Folio                   | 8 1/2 x 13 in        |
    /// | 15       | Quarto                  | 215 x 275 mm         |
    /// | 16       | ---                     | 10x14 in             |
    /// | 17       | ---                     | 11x17 in             |
    /// | 18       | Note                    | 8 1/2 x 11 in        |
    /// | 19       | Envelope 9              | 3 7/8 x 8 7/8        |
    /// | 20       | Envelope 10             | 4 1/8 x 9 1/2        |
    /// | 21       | Envelope 11             | 4 1/2 x 10 3/8       |
    /// | 22       | Envelope 12             | 4 3/4 x 11           |
    /// | 23       | Envelope 14             | 5 x 11 1/2           |
    /// | 24       | C size sheet            | ---                  |
    /// | 25       | D size sheet            | ---                  |
    /// | 26       | E size sheet            | ---                  |
    /// | 27       | Envelope DL             | 110 x 220 mm         |
    /// | 28       | Envelope C3             | 324 x 458 mm         |
    /// | 29       | Envelope C4             | 229 x 324 mm         |
    /// | 30       | Envelope C5             | 162 x 229 mm         |
    /// | 31       | Envelope C6             | 114 x 162 mm         |
    /// | 32       | Envelope C65            | 114 x 229 mm         |
    /// | 33       | Envelope B4             | 250 x 353 mm         |
    /// | 34       | Envelope B5             | 176 x 250 mm         |
    /// | 35       | Envelope B6             | 176 x 125 mm         |
    /// | 36       | Envelope                | 110 x 230 mm         |
    /// | 37       | Monarch                 | 3.875 x 7.5 in       |
    /// | 38       | Envelope                | 3 5/8 x 6 1/2 in     |
    /// | 39       | Fanfold                 | 14 7/8 x 11 in       |
    /// | 40       | German Std Fanfold      | 8 1/2 x 12 in        |
    /// | 41       | German Legal Fanfold    | 8 1/2 x 13 in        |
    ///
    /// Note, it is likely that not all of these paper types will be available
    /// to the end user since it will depend on the paper formats that the
    /// user's printer supports. Therefore, it is best to stick to standard
    /// paper types of 1 for US Letter and 9 for A4.
    ///
    /// If you do not specify a paper type the worksheet will print using the
    /// printer's default paper style.
    ///
    /// # Parameters
    ///
    /// - `paper_size`: The paper size index from the list above .
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet paper size/type for
    /// the printed output.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_paper.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set the printer paper size.
    ///     worksheet.set_paper_size(9); // A4 paper size.
    ///
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn set_paper_size(&mut self, paper_size: u8) -> &mut Worksheet {
        self.paper_size = paper_size;
        self.page_setup_changed = true;
        self
    }

    /// Set the order in which pages are printed.
    ///
    /// The `set_page_order()` method is used to change the default print
    /// direction. This is referred to by Excel as the sheet "page order":
    ///
    /// The default page order is shown below for a worksheet that extends over
    /// 4 pages. The order is called "down then over":
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_page_order.png">
    ///
    /// However, by using `set_page_order(false)` the print order will be
    /// changed to "over then down".
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. Set `true` to get "Down, then
    ///   over" (the default) and `false` to get "Over, then down".
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet printed page
    /// order.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_page_order.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set the page print to "over then down"
    ///     worksheet.set_page_order(false);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn set_page_order(&mut self, enable: bool) -> &mut Worksheet {
        self.default_page_order = enable;

        if !enable {
            self.page_setup_changed = true;
        }
        self
    }

    /// Set the page orientation to landscape.
    ///
    /// The `set_landscape()` method is used to set the orientation of a
    /// worksheet's printed page to landscape.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet page orientation to
    /// landscape.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_landscape.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.set_landscape();
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_landscape(&mut self) -> &mut Worksheet {
        self.is_portrait = false;
        self.page_setup_changed = true;
        self
    }

    /// Set the page orientation to portrait.
    ///
    ///  This `set_portrait()` method  is used to set the orientation of a
    ///  worksheet's printed page to portrait. The default worksheet orientation
    ///  is portrait, so this function is rarely required.
    ///
    pub fn set_portrait(&mut self) -> &mut Worksheet {
        self.is_portrait = true;
        self.page_setup_changed = true;
        self
    }

    /// Set the page view mode to normal layout.
    ///
    /// This method is used to display the worksheet in “View -> Normal”
    /// mode. This is the default.
    ///
    pub fn set_view_normal(&mut self) -> &mut Worksheet {
        self.page_view = PageView::Normal;
        self
    }

    /// Set the page view mode to page layout.
    ///
    /// This method is used to display the worksheet in “View -> Page Layout”
    /// mode.
    ///
    pub fn set_view_page_layout(&mut self) -> &mut Worksheet {
        self.page_view = PageView::PageLayout;
        self.page_setup_changed = true;
        self
    }

    /// Set the page view mode to page break preview.
    ///
    /// This method is used to display the worksheet in “View -> Page Break
    /// Preview” mode.
    ///
    pub fn set_view_page_break_preview(&mut self) -> &mut Worksheet {
        self.page_view = PageView::PageBreaks;
        self.page_setup_changed = true;
        self
    }

    /// Set the horizontal page breaks on a worksheet.
    ///
    /// The `set_page_breaks()` method adds horizontal page breaks to a
    /// worksheet. A page break causes all the data that follows it to be
    /// printed on the next page. Horizontal page breaks act between rows.
    ///
    /// # Parameters
    ///
    /// - `breaks`: A list of one or more row numbers where the page breaks
    ///   occur. To create a page break between rows 20 and 21 you must specify
    ///   the break at row 21. However in zero index notation this is actually
    ///   row 20. So you can pretend for a small while that you are using 1
    ///   index notation.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::ParameterError`] - The number of page breaks exceeds
    ///   Excel's limit of 1023 page breaks.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting page breaks for a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_page_breaks.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(100, 100, "Test")?;
    /// #
    ///     // Set a page break at rows 20, 40 and 60.
    ///     worksheet.set_page_breaks(&[20, 40, 60])?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_page_breaks.png">
    ///
    pub fn set_page_breaks(&mut self, breaks: &[RowNum]) -> Result<&mut Worksheet, XlsxError> {
        // Ignore empty input.
        if breaks.is_empty() {
            return Ok(self);
        }

        // Sort list and remove any duplicates and 0.
        let breaks = Self::process_pagebreaks(breaks)?;

        // Check max break value is within Excel column limit.
        if *breaks.last().unwrap() >= ROW_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        self.horizontal_breaks = breaks;

        Ok(self)
    }

    /// Set the vertical page breaks on a worksheet.
    ///
    /// The `set_vertical_page_breaks()` method adds vertical page breaks to a
    /// worksheet. This is much less common than the
    /// [`Worksheet::set_page_breaks()`] method shown above.
    ///
    /// # Parameters
    ///
    /// - `breaks`: A list of one or more column numbers where the page breaks
    ///   occur.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::ParameterError`] - The number of page breaks exceeds
    ///   Excel's limit of 1023 page breaks.
    ///
    pub fn set_vertical_page_breaks(
        &mut self,
        breaks: &[u32],
    ) -> Result<&mut Worksheet, XlsxError> {
        // Ignore empty input.
        if breaks.is_empty() {
            return Ok(self);
        }

        // Sort list and remove any duplicates and 0.
        let breaks = Self::process_pagebreaks(breaks)?;

        // Check max break value is within Excel col limit.
        if *breaks.last().unwrap() >= u32::from(COL_MAX) {
            return Err(XlsxError::RowColumnLimitError);
        }

        self.vertical_breaks = breaks;

        Ok(self)
    }

    /// Set the worksheet zoom factor.
    ///
    /// Set the worksheet zoom factor in the range 10 <= zoom <= 400.
    ///
    /// The default zoom level is 100. The `set_zoom()` method does not affect
    /// the scale of the printed page in Excel. For that you should use
    /// [`Worksheet::set_print_scale()`].
    ///
    /// # Parameters
    ///
    /// - `zoom`: The worksheet zoom level.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet zoom level.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_zoom.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///     worksheet.set_zoom(200);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_zoom.png">
    ///
    pub fn set_zoom(&mut self, zoom: u16) -> &mut Worksheet {
        if !(10..=400).contains(&zoom) {
            eprintln!("Zoom factor {zoom} outside Excel range: 10 <= zoom <= 400.");
            return self;
        }

        self.zoom = zoom;
        self
    }

    /// Freeze panes in a worksheet.
    ///
    /// The `set_freeze_panes()` method can be used to divide a worksheet into
    /// horizontal or vertical regions known as panes and to “freeze” these
    /// panes so that the splitter bars are not visible.
    ///
    /// As with Excel the split is to the top and left of the cell. So to freeze
    /// the top row and leftmost column you would use `(1, 1)` (zero-indexed).
    /// Also, you can set one of the row and col parameters as 0 if you do not
    /// want either the vertical or horizontal split. See the example below.
    ///
    /// In Excel it is also possible to set "split" panes without freezing them.
    /// That feature isn't currently supported by `rust_xlsxwriter`.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet panes.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_freeze_panes.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let mut worksheet1 = Worksheet::new();
    /// #     let mut worksheet2 = Worksheet::new();
    /// #     let mut worksheet3 = Worksheet::new();
    /// #
    /// #     worksheet1.write_string(0, 0, "Scroll down")?;
    /// #     worksheet2.write_string(0, 0, "Scroll across")?;
    /// #     worksheet3.write_string(0, 0, "Scroll down or across")?;
    /// #
    ///     // Freeze the top row only.
    ///     worksheet1.set_freeze_panes(1, 0)?;
    ///
    ///     // Freeze the leftmost column only.
    ///     worksheet2.set_freeze_panes(0, 1)?;
    ///
    ///     // Freeze the top row and leftmost column.
    ///     worksheet3.set_freeze_panes(1, 1)?;
    /// #
    /// #     workbook.push_worksheet(worksheet1);
    /// #     workbook.push_worksheet(worksheet2);
    /// #     workbook.push_worksheet(worksheet3);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_freeze_panes.png">
    ///
    pub fn set_freeze_panes(
        &mut self,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        self.panes.freeze_cell = (row, col);
        Ok(self)
    }

    /// Set the top most cell in the scrolling area of a freeze pane.
    ///
    /// This method is used in conjunction with the
    /// [`Worksheet::set_freeze_panes()`] method to set the top most visible
    /// cell in the scrolling range. For example you may want to freeze the top
    /// row but have the worksheet pre-scrolled so that cell `A20` is visible in
    /// the scrolled area. See the example below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet panes and also
    /// setting the topmost visible cell in the scrolled area.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_freeze_panes_top_cell.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write_string(0, 0, "Scroll down")?;
    /// #
    ///     // Freeze the top row only.
    ///     worksheet.set_freeze_panes(1, 0)?;
    ///
    ///     // Pre-scroll to the row 20.
    ///     worksheet.set_freeze_panes_top_cell(19, 0)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_freeze_panes_top_cell.png">
    ///
    pub fn set_freeze_panes_top_cell(
        &mut self,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        self.panes.top_cell = (row, col);
        Ok(self)
    }

    /// Set the printed page header caption.
    ///
    /// The `set_header()` method can be used to set the header for a worksheet.
    ///
    /// Headers and footers are generated using a string which is a combination
    /// of plain text and optional control characters.
    ///
    /// The available control characters are:
    ///
    /// | Control              | Category      | Description           |
    /// | -------------------- | ------------- | --------------------- |
    /// | `&L`                 | Alignment     | Left                  |
    /// | `&C`                 |               | Center                |
    /// | `&R`                 |               | Right                 |
    /// | `&[Page]`  or `&P`   | Information   | Page number           |
    /// | `&[Pages]` or `&N`   |               | Total number of pages |
    /// | `&[Date]`  or `&D`   |               | Date                  |
    /// | `&[Time]`  or `&T`   |               | Time                  |
    /// | `&[File]`  or `&F`   |               | File name             |
    /// | `&[Tab]`   or `&A`   |               | Worksheet name        |
    /// | `&[Path]`  or `&Z`   |               | Workbook path         |
    /// | `&fontsize`          | Font          | Font size             |
    /// | `&"font,style"`      |               | Font name and style   |
    /// | `&U`                 |               | Single underline      |
    /// | `&E`                 |               | Double underline      |
    /// | `&S`                 |               | Strikethrough         |
    /// | `&X`                 |               | Superscript           |
    /// | `&Y`                 |               | Subscript             |
    /// | `&[Picture]` or `&G` | Images        | Picture/image         |
    /// | `&&`                 | Miscellaneous | Literal ampersand &   |
    ///
    /// Some of the placeholder variables have a long version like `&[Page]` and
    /// a short version like `&P`. The longer version is displayed in the Excel
    /// interface but the shorter version is the way that it is stored in the
    /// file format. Either version is okay since `rust_xlsxwriter` will
    /// translate as required.
    ///
    /// Headers and footers have 3 edit areas to the left, center and right.
    /// Text can be aligned to these areas by prefixing the text with the
    /// control characters `&L`, `&C` and `&R`.
    ///
    /// For example:
    ///
    /// ```text
    /// worksheet.set_header("&LHello");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// | Hello                                                         |
    /// |                                                               |
    ///
    /// worksheet.set_header("&CHello");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                          Hello                                |
    /// |                                                               |
    ///
    /// worksheet.set_header("&RHello");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                                                         Hello |
    /// |                                                               |
    /// ```
    ///
    /// You can also have text in each of the alignment areas:
    ///
    /// ```text
    /// worksheet.set_header("&LCiao&CBello&RCielo");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// | Ciao                     Bello                          Cielo |
    /// |                                                               |
    /// ```
    ///
    /// The information control characters act as variables/templates that Excel
    /// will update/expand as the workbook or worksheet changes.
    ///
    /// ```text
    /// worksheet.set_header("&CPage &[Page] of &[Pages]");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                        Page 1 of 6                            |
    /// |                                                               |
    /// ```
    ///
    /// Times and dates are in the user's default format:
    ///
    /// ```text
    /// worksheet.set_header("&CUpdated at &[Time]");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                    Updated at 12:30 PM                        |
    /// |                                                               |
    /// ```
    ///
    /// To include a single literal ampersand `&` in a header or footer you
    /// should use a double ampersand `&&`:
    ///
    /// ```text
    /// worksheet.set_header("&CCuriouser && Curiouser - Attorneys at Law");
    ///
    ///  ---------------------------------------------------------------
    /// |                                                               |
    /// |                   Curiouser & Curiouser                       |
    /// |                                                               |
    /// ```
    ///
    /// You can specify the font size of a section of the text by prefixing it
    /// with the control character `&n` where `n` is the font size:
    ///
    /// ```text
    /// worksheet1.set_header("&C&30Hello Big");
    /// worksheet2.set_header("&C&10Hello Small");
    /// ```
    ///
    /// You can specify the font of a section of the text by prefixing it with
    /// the control sequence `&"font,style"` where `fontname` is a font name
    /// such as Windows font descriptions: "Regular", "Italic", "Bold" or "Bold
    /// Italic": "Courier New" or "Times New Roman" and `style` is one of the
    /// standard Windows font descriptions like “Regular”, “Italic”, “Bold” or
    /// “Bold Italic”:
    ///
    /// ```text
    /// worksheet1.set_header(r#"&C&"Courier New,Italic"Hello"#);
    /// worksheet2.set_header(r#"&C&"Courier New,Bold Italic"Hello"#);
    /// worksheet3.set_header(r#"&C&"Times New Roman,Regular"Hello"#);
    /// ```
    ///
    /// It is possible to combine all of these features together to create
    /// complex headers and footers. If you set up a complex header in Excel you
    /// can transfer it to `rust_xlsxwriter` by inspecting the string in the
    /// Excel file. For example the following shows how unzip and grep the Excel
    /// XML sub-files on a Linux system. The example uses libxml's xmllint to
    /// format the XML for clarity:
    ///
    /// ```text
    /// $ unzip myfile.xlsm -d myfile
    /// $ xmllint --format `find myfile -name "*.xml" | xargs` | \
    ///     egrep "Header|Footer" | sed 's/&amp;/\&/g'
    ///
    ///  <headerFooter scaleWithDoc="0">
    ///    <oddHeader>&L&P</oddHeader>
    ///  </headerFooter>
    /// ```
    ///
    /// Note: Excel requires that the header or footer string be less than 256
    /// characters, including the control characters. Strings longer than this
    /// will not be written, and a warning will be output.
    ///
    /// # Parameters
    ///
    /// - `header`: The header string with optional control characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet header.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_header.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.set_header("&CPage &P of &N");
    /// #
    /// #     worksheet.write_string(0, 0, "Hello")?;
    /// #     worksheet.write_string(200, 0, "Hello")?;
    /// #     worksheet.set_view_page_layout();
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_header.png">
    ///
    pub fn set_header(&mut self, header: impl Into<String>) -> &mut Worksheet {
        let header = header.into();
        let header_expanded = header
            .replace("&[Tab]", "&A")
            .replace("&[Date]", "&D")
            .replace("&[File]", "&F")
            .replace("&[Page]", "&P")
            .replace("&[Path]", "&Z")
            .replace("&[Time]", "&T")
            .replace("&[Pages]", "&N")
            .replace("&[Picture]", "&G");

        if header_expanded.chars().count() > 255 {
            eprintln!("Header string exceeds Excel's limit of 255 characters.");
            return self;
        }

        self.header = header;
        self.page_setup_changed = true;
        self.head_footer_changed = true;
        self
    }

    /// Set the printed page footer caption.
    ///
    /// The `set_footer()` method can be used to set the footer for a worksheet.
    ///
    /// See the documentation for [`Worksheet::set_header()`] for more details
    /// on the syntax of the header/footer string.
    ///
    /// # Parameters
    ///
    /// - `footer`: The footer string with optional control characters.
    ///
    pub fn set_footer(&mut self, footer: impl Into<String>) -> &mut Worksheet {
        let footer = footer.into();
        let footer_expanded = footer
            .replace("&[Tab]", "&A")
            .replace("&[Date]", "&D")
            .replace("&[File]", "&F")
            .replace("&[Page]", "&P")
            .replace("&[Path]", "&Z")
            .replace("&[Time]", "&T")
            .replace("&[Pages]", "&N")
            .replace("&[Picture]", "&G");

        if footer_expanded.chars().count() > 255 {
            eprintln!("Footer string exceeds Excel's limit of 255 characters.");
            return self;
        }

        self.footer = footer;
        self.page_setup_changed = true;
        self.head_footer_changed = true;
        self
    }

    /// Insert an image in a worksheet header.
    ///
    /// Insert an image in a worksheet header in one of the 3 sections supported
    /// by Excel: Left, Center and Right. This needs to be preceded by a call to
    /// [`Worksheet::set_header()`] where a corresponding `&[Picture]` element
    /// is added to the header formatting string such as `"&L&[Picture]"`.
    ///
    /// # Parameters
    ///
    /// - `position`: The image position as defined by the
    ///   [`HeaderImagePosition`] enum.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::ParameterError`] - Parameter error if there isn't a
    ///   corresponding `&[Picture]`/`&[G]` variable in the header string.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a header image to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_header_image.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError, HeaderImagePosition};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Scale the image so it fits in the header.
    ///       let image = Image::new("examples/rust_logo.png")?
    ///           .set_scale_height(0.5)
    ///           .set_scale_width(0.5);
    /// #
    ///     // Insert the watermark image in the header.
    ///     worksheet.set_header("&C&[Picture]");
    ///     worksheet.set_header_image(&image, HeaderImagePosition::Center)?;
    /// #
    /// #     // Increase the top margin to 1.2 for clarity. The -1.0 values are ignored.
    /// #     worksheet.set_margins(-1.0, -1.0, 1.2, -1.0, -1.0, -1.0);
    /// #
    /// #     // Set Page View mode so the watermark is visible.
    /// #     worksheet.set_view_page_layout();
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_header_image.png">
    ///
    /// An example of adding a worksheet watermark image using the
    /// `rust_xlsxwriter` library. This is based on the method of putting an
    /// image in the worksheet header as suggested in the [Microsoft
    /// documentation].
    ///
    /// [Microsoft documentation]:
    ///     https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
    ///
    /// ```
    /// # // This code is available in examples/app_watermark.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError, HeaderImagePosition};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     let image = Image::new("examples/watermark.png")?;
    ///
    ///     // Insert the watermark image in the header.
    ///     worksheet.set_header("&C&[Picture]");
    ///     worksheet.set_header_image(&image, HeaderImagePosition::Center)?;
    /// #
    /// #     // Set Page View mode so the watermark is visible.
    /// #     worksheet.set_view_page_layout();
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("watermark.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/app_watermark.png">
    ///
    pub fn set_header_image(
        &mut self,
        image: &Image,
        position: HeaderImagePosition,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check that there is a matching  &[Picture]/&[G] variable in the
        // header string.
        if !Self::verify_header_footer_image(&self.header, &position) {
            let error = format!(
                "No &[Picture] or &[G] variable in header string: '{}' for position = '{:?}'",
                self.header, position
            );
            return Err(XlsxError::ParameterError(error));
        }

        let mut image = image.clone();
        image.header_position = position.clone();
        image.is_header = true;
        self.header_footer_images[position as usize] = Some(image);

        Ok(self)
    }

    /// Insert an image in a worksheet footer.
    ///
    /// See the documentation for [`Worksheet::set_header_image()`] for more
    /// details.
    ///
    /// # Parameters
    ///
    /// - `position`: The image position as defined by the
    ///   [`HeaderImagePosition`] enum.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::ParameterError`] - Parameter error if there isn't a
    ///   corresponding `&[Picture]`/`&[G]` variable in the header string.
    ///
    pub fn set_footer_image(
        &mut self,
        image: &Image,
        position: HeaderImagePosition,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check that there is a matching  &[Picture]/&[G] variable in the
        // footer string.
        if !Self::verify_header_footer_image(&self.footer, &position) {
            let error = format!(
                "No &[Picture] or &[G] variable in footer string: '{}' for position = '{:?}'",
                self.footer, position
            );
            return Err(XlsxError::ParameterError(error));
        }

        let mut image = image.clone();
        image.header_position = position.clone();
        image.is_header = false;
        self.header_footer_images[3 + position as usize] = Some(image);

        Ok(self)
    }

    /// Set the page setup option to scale the header/footer with the document.
    ///
    /// This option determines whether the headers and footers use the same
    /// scaling as the worksheet. This defaults to "on" in Excel.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Header/Footer](../worksheet/index.html#page-setup---headerfooter).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn set_header_footer_scale_with_doc(&mut self, enable: bool) -> &mut Worksheet {
        self.header_footer_scale_with_doc = enable;

        if !enable {
            self.page_setup_changed = true;
            self.head_footer_changed = true;
        }

        self
    }

    /// Set the page setup option to align the header/footer with the margins.
    ///
    /// This option determines whether the headers and footers align with the
    /// left and right margins of the worksheet. This defaults to "on" in Excel.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Header/Footer](../worksheet/index.html#page-setup---headerfooter).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default. S
    pub fn set_header_footer_align_with_page(&mut self, enable: bool) -> &mut Worksheet {
        self.header_footer_align_with_page = enable;

        if !enable {
            self.page_setup_changed = true;
            self.head_footer_changed = true;
        }
        self
    }

    /// Set the page margins.
    ///
    /// The `set_margins()` method is used to set the margins of the worksheet
    /// when it is printed. The units are in inches. Specifying `-1.0` for any
    /// parameter will give the default Excel value. The defaults are shown
    /// below.
    ///
    /// # Parameters
    ///
    /// - `left`: Left margin in inches. Excel default is 0.7.
    /// - `right`: Right margin in inches. Excel default is 0.7.
    /// - `top`: Top margin in inches. Excel default is 0.75.
    /// - `bottom`: Bottom margin in inches. Excel default is 0.75.
    /// - `header`: Header margin in inches. Excel default is 0.3.
    /// - `footer`: Footer margin in inches. Excel default is 0.3.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the worksheet margins.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_margins.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.set_margins(1.0, 1.25, 1.5, 1.75, 0.75, 0.25);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_margins.png">
    ///
    pub fn set_margins(
        &mut self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> &mut Worksheet {
        if left >= 0.0 {
            self.margin_left = left;
            self.page_setup_changed = true;
        }
        if right >= 0.0 {
            self.margin_right = right;
            self.page_setup_changed = true;
        }
        if top >= 0.0 {
            self.margin_top = top;
            self.page_setup_changed = true;
        }
        if bottom >= 0.0 {
            self.margin_bottom = bottom;
            self.page_setup_changed = true;
        }
        if header >= 0.0 {
            self.margin_header = header;
            self.page_setup_changed = true;
        }
        if footer >= 0.0 {
            self.margin_footer = footer;
            self.page_setup_changed = true;
        }

        self
    }

    /// Set the first page number when printing.
    ///
    /// The `set_print_first_page_number()` method is used to set the page
    /// number of the first page when the worksheet is printed out. This option
    /// will only have and effect if you have a header/footer with the `&[Page]`
    /// control character, see [`Worksheet::set_header()`].
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `page_number`: The page number of the first printed page.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the page number on the
    /// printed page.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_print_first_page_number.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.set_header("&CPage &P of &N");
    ///     worksheet.set_print_first_page_number(2);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn set_print_first_page_number(&mut self, page_number: u16) -> &mut Worksheet {
        self.first_page_number = page_number;
        self.page_setup_changed = true;
        self
    }

    /// Set the page setup option to set the print scale.
    ///
    /// Set the scale factor of the printed page, in the range 10 <= scale <=
    /// 400.
    ///
    /// The default scale factor is 100. The `set_print_scale()` method
    /// does not affect the scale of the visible page in Excel. For that you
    /// should use [`Worksheet::set_zoom()`].
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `scale`: The print scale factor.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the scale of the worksheet page
    /// when printed.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_print_scale.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Scale the printed worksheet to 50%.
    ///     worksheet.set_print_scale(50);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_print_scale(&mut self, scale: u16) -> &mut Worksheet {
        if !(10..=400).contains(&scale) {
            eprintln!("Scale factor {scale} outside Excel range: 10 <= zoom <= 400.");
            return self;
        }

        // This property is mutually exclusive with fit to page.
        self.fit_to_page = false;

        self.print_scale = scale;
        self.page_setup_changed = true;
        self
    }

    /// Fit the printed area to a specific number of pages both vertically and
    /// horizontally.
    ///
    /// The `set_print_fit_to_pages()` method is used to fit the printed area to
    /// a specific number of pages both vertically and horizontally. If the
    /// printed area exceeds the specified number of pages it will be scaled
    /// down to fit. This ensures that the printed area will always appear on
    /// the specified number of pages even if the page size or margins change:
    ///
    /// ```text
    ///     worksheet1.set_print_fit_to_pages(1, 1); // Fit to 1x1 pages.
    ///     worksheet2.set_print_fit_to_pages(2, 1); // Fit to 2x1 pages.
    ///     worksheet3.set_print_fit_to_pages(1, 2); // Fit to 1x2 pages.
    /// ```
    ///
    /// The print area can be defined using the `set_print_area()` method.
    ///
    /// A common requirement is to fit the printed output to `n` pages wide but
    /// have the height be as long as necessary. To achieve this set the
    /// `height` to 0, see the example below.
    ///
    /// **Notes**:
    ///
    /// - The `set_print_fit_to_pages()` will override any manual page breaks
    ///   that are defined in the worksheet.
    ///
    /// - When using `set_print_fit_to_pages()` it may also be required to set
    ///   the printer paper size using [`Worksheet::set_paper_size()`] or else
    ///   Excel will default to "US Letter".
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `width`: Number of pages horizontally.
    /// - `height`: Number of pages vertically.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the scale of the worksheet to
    /// fit a defined number of pages vertically and horizontally. This example
    /// shows a common use case which is to fit the printed output to 1 page
    /// wide but have the height be as long as necessary.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_print_fit_to_pages.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set the printed output to fit 1 page wide and as long as necessary.
    ///     worksheet.set_print_fit_to_pages(1, 0);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_print_fit_to_pages.png">
    ///
    pub fn set_print_fit_to_pages(&mut self, width: u16, height: u16) -> &mut Worksheet {
        self.fit_width = width;
        self.fit_height = height;

        // This property is mutually exclusive with print scale.
        self.print_scale = 100;

        self.fit_to_page = true;
        self.page_setup_changed = true;
        self
    }

    /// Center the printed page horizontally.
    ///
    /// Center the worksheet data horizontally between the margins on the
    /// printed page
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Margins](../worksheet/index.html#page-setup---margins).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_print_center_horizontally(&mut self, enable: bool) -> &mut Worksheet {
        self.center_horizontally = enable;

        if enable {
            self.print_options_changed = true;
            self.page_setup_changed = true;
        }
        self
    }

    /// Center the printed page vertically.
    ///
    /// Center the worksheet data vertically between the margins on the printed
    /// page
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Margins](../worksheet/index.html#page-setup---margins).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_print_center_vertically(&mut self, enable: bool) -> &mut Worksheet {
        self.center_vertically = enable;

        if enable {
            self.print_options_changed = true;
            self.page_setup_changed = true;
        }
        self
    }

    /// Set the option to turn on/off the screen gridlines.
    ///
    /// The `set_screen_gridlines()` method is use to turn on/off gridlines on
    /// displayed worksheet. It is on by default.
    ///
    /// To turn on/off the printed gridlines see the
    /// [`Worksheet::set_print_gridlines()`] method below.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates turn off the worksheet worksheet screen
    /// gridlines.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_screen_gridlines.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     worksheet.write(0, 0, "Hello")?;
    ///
    ///     // Turn off the screen gridlines.
    ///     worksheet.set_screen_gridlines(false);
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_set_screen_gridlines.png">
    ///
    pub fn set_screen_gridlines(&mut self, enable: bool) -> &mut Worksheet {
        self.screen_gridlines = enable;

        self
    }

    /// Set the page setup option to turn on printed gridlines.
    ///
    /// The `set_print_gridlines()` method is use to turn on/off gridlines on
    /// the printed pages. It is off by default.
    ///
    /// To turn on/off the screen gridlines see the
    /// [`Worksheet::set_screen_gridlines()`] method above.
    ///
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_print_gridlines(&mut self, enable: bool) -> &mut Worksheet {
        self.print_gridlines = enable;

        if enable {
            self.print_options_changed = true;
            self.page_setup_changed = true;
        }
        self
    }

    /// Set the page setup option to print in black and white.
    ///
    /// This `set_print_black_and_white()` method can be used to force printing
    /// in black and white only. It is off by default.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_print_black_and_white(&mut self, enable: bool) -> &mut Worksheet {
        self.print_black_and_white = enable;

        if enable {
            self.page_setup_changed = true;
        }
        self
    }

    /// Set the page setup option to print in draft quality.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_print_draft(&mut self, enable: bool) -> &mut Worksheet {
        self.print_draft = enable;

        if enable {
            self.page_setup_changed = true;
        }
        self
    }

    /// Set the page setup option to print the row and column headers on the
    /// printed page.
    ///
    /// The `set_print_headings()` method turns on the row and column headers
    /// when printing a worksheet. This option is off by default.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_print_headings(&mut self, enable: bool) -> &mut Worksheet {
        self.print_headings = enable;

        if enable {
            self.print_options_changed = true;
            self.page_setup_changed = true;
        }
        self
    }

    /// Set the print area for the worksheet.
    ///
    /// This method is used to specify the area of the worksheet that will be
    /// printed.
    ///
    /// In order to specify an entire row or column range such as `1:20` or
    /// `A:H` you must specify the corresponding maximum column or row range.
    /// For example:
    ///
    /// - `(0, 0, 31, 16_383) == 1:32`.
    /// - `(0, 0, 1_048_575, 12) == A:M`.
    ///
    /// In these examples 16,383 is the maximum column and 1,048,575 is the
    /// maximum row (zero indexed).
    ///
    /// See also the example below and the documentation on
    /// [Worksheet Page Setup - Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row or column is larger
    ///   than the last row or column.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the print area for several
    /// worksheets.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_print_area.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = workbook.add_worksheet();
    ///     // Set the print area to "A1:M32"
    ///     worksheet1.set_print_area(0, 0, 31, 12)?;
    ///
    ///     let worksheet2 = workbook.add_worksheet();
    ///     // Set the print area to "1:32"
    ///     worksheet2.set_print_area(0, 0, 31, 16_383)?;
    ///
    ///     let worksheet3 = workbook.add_worksheet();
    ///     // Set the print area to "A:M"
    ///     worksheet3.set_print_area(0, 0, 1_048_575, 12)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file, page setup dialog for worksheet1:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_print_area.png">
    ///
    pub fn set_print_area(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // The print range is the entire worksheet, therefore it is the same
        // as the default, so we can ignore it.
        if first_row == 0 && first_col == 0 && last_row == ROW_MAX - 1 && last_col == COL_MAX - 1 {
            return Ok(self);
        }

        // Store the defined name information.
        self.print_area_defined_name.in_use = true;
        self.print_area_defined_name.name_type = DefinedNameType::PrintArea;
        self.print_area_defined_name.first_row = first_row;
        self.print_area_defined_name.first_col = first_col;
        self.print_area_defined_name.last_row = last_row;
        self.print_area_defined_name.last_col = last_col;

        self.page_setup_changed = true;
        Ok(self)
    }

    /// Set the number of rows to repeat at the top of each printed page.
    ///
    /// For large Excel documents it is often desirable to have the first row or
    /// rows of the worksheet print out at the top of each page.
    ///
    /// See the example below and the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (Zero indexed.)
    /// - `last_row`: The last row of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row greater than the last
    ///   row.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the rows to repeat on each
    /// printed page.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_repeat_rows.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = workbook.add_worksheet();
    ///     // Repeat the first row in the printed output.
    ///     worksheet1.set_repeat_rows(0, 0)?;
    ///
    ///     let worksheet2 = workbook.add_worksheet();
    ///     // Repeat the first 2 rows in the printed output.
    ///     worksheet2.set_repeat_rows(0, 1)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file, page setup dialog for worksheet2:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_repeat_rows.png">
    ///
    pub fn set_repeat_rows(
        &mut self,
        first_row: RowNum,
        last_row: RowNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows are in the allowed range.
        if !self.check_dimensions_only(first_row, 0) || !self.check_dimensions_only(last_row, 0) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Store the range data.
        self.repeat_row_cols_defined_name.in_use = true;
        self.repeat_row_cols_defined_name.name_type = DefinedNameType::PrintTitles;
        self.repeat_row_cols_defined_name.first_row = first_row;
        self.repeat_row_cols_defined_name.last_row = last_row;

        self.page_setup_changed = true;
        Ok(self)
    }

    /// Set the columns to repeat at the left hand side of each printed page.
    ///
    /// For large Excel documents it is often desirable to have the first column
    /// or columns of the worksheet print out at the left hand side of each
    /// page.
    ///
    /// See the example below and the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. (Zero indexed.)
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row or column is larger
    ///   than the last row or column.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting the columns to repeat on each
    /// printed page.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_repeat_columns.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    ///     let worksheet1 = workbook.add_worksheet();
    ///     // Repeat the first column in the printed output.
    ///     worksheet1.set_repeat_columns(0, 0)?;
    ///
    ///     let worksheet2 = workbook.add_worksheet();
    ///     // Repeat the first 2 columns in the printed output.
    ///     worksheet2.set_repeat_columns(0, 1)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file, page setup dialog for worksheet2:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_repeat_columns.png">
    ///
    pub fn set_repeat_columns(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check columns are in the allowed range.
        if !self.check_dimensions_only(0, first_col) || !self.check_dimensions_only(0, last_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Store the defined name information.
        self.repeat_row_cols_defined_name.in_use = true;
        self.repeat_row_cols_defined_name.name_type = DefinedNameType::PrintTitles;
        self.repeat_row_cols_defined_name.first_col = first_col;
        self.repeat_row_cols_defined_name.last_col = last_col;

        self.page_setup_changed = true;
        Ok(self)
    }

    /// Autofit the worksheet column widths to the widest data in the column,
    /// approximately.
    ///
    /// Excel autofits columns at runtime when it has access to all of the
    /// required worksheet information as well as the Windows functions for
    /// calculating display areas based on fonts and formatting.
    ///
    /// The `rust_xlsxwriter` library doesn't have access to these Windows
    /// functions so it simulates autofit by calculating string widths based on
    /// metrics taken from Excel.
    ///
    /// This isn't perfect but for most cases it should be sufficient and
    /// indistinguishable from the output of Excel. However there are some
    /// limitations to be aware of when using this method:
    ///
    /// - It is based on the default Excel font type and size of Calibri 11. It
    ///   will not give accurate results for other fonts or font sizes.
    /// - It doesn't take formatting of numbers or dates account, although this
    ///   may be addressed in a later version.
    /// - Autofit is a relatively expensive operation since it performs a
    ///   calculation for all the populated cells in a worksheet. See the note
    ///   on performance below.
    ///
    /// For cases that don't match your desired output you can set explicit
    /// column widths via [`Worksheet::set_column_width()`] or
    /// [`Worksheet::set_column_width_pixels()`]. The `autofit()` method ignores
    /// columns that have already been explicitly set if the width is greater
    /// than the calculated autofit width. Alternatively, setting the column
    /// width explicitly after calling `autofit()` will override the autofit
    /// value. See also [`Worksheet::autofit_to_max_width()`] below.
    ///
    /// **Performance**: By default `autofit()` performs a length calculation
    /// for each populated cell in a worksheet. For very large worksheets this
    /// could be slow. However, it is possible to mitigate this by calling
    /// `autofit()` after writing the first 100 or 200 rows. This will produce a
    /// reasonably accurate autofit for the first visible page of data without
    /// incurring the performance penalty of autofitting thousands of
    /// non-visible rows.
    ///
    /// # Examples
    ///
    /// The following example demonstrates auto-fitting the worksheet column
    /// widths based on the data in the columns. See all the [Autofitting
    /// Columns] example in the user guide/examples directory.
    ///
    /// [Autofitting Columns]:
    ///     https://rustxlsxwriter.github.io/examples/autofit.html
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_autofit.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add some data to the worksheet.
    ///     worksheet.write_string(0, 0, "Hello")?;
    ///     worksheet.write_string(0, 1, "Hello")?;
    ///     worksheet.write_string(1, 1, "Hello World")?;
    ///     worksheet.write_number(0, 2, 123)?;
    ///     worksheet.write_number(0, 3, 123456)?;
    ///
    ///     // Autofit the columns.
    ///     worksheet.autofit();
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_autofit.png">
    ///
    ///
    ///
    pub fn autofit(&mut self) -> &mut Worksheet {
        self.autofit_worksheet(MAX_AUTOFIT_WIDTH_PIXELS)
    }

    /// Autofit the worksheet columns up to a maximum width.
    ///
    /// The [`Worksheet::autofit()`] method above simulates Excel's column
    /// autofit. On undesirable side-effect of this is that Excel autofits very
    /// long strings up to limit of 255 characters/1790 pixels. This is often
    /// too wide to display on a single screen at normal zoom. As such the
    /// `autofit_to_max_width()` method is provided to enable a smaller upper
    /// limit for autofitting long strings. A value of 300 pixels is recommended
    /// as a good compromise between column width and readability.
    ///
    /// # Parameters
    ///
    /// - `max_autofit_width`: The maximum column width, in pixels, to use for
    ///   autofitting.
    ///
    ///
    pub fn autofit_to_max_width(&mut self, max_autofit_width: u16) -> &mut Worksheet {
        let max_autofit_width = std::cmp::min(max_autofit_width, MAX_AUTOFIT_WIDTH_PIXELS);
        self.autofit_worksheet(max_autofit_width)
    }

    /// Set the worksheet name used in VBA macros.
    ///
    /// This method can be used to set the VBA name for the worksheet. This is
    /// sometimes required when a VBA macro included via
    /// [`Workbook::add_vba_project()`](crate::Workbook::add_vba_project())
    /// makes reference to the worksheet with a name other than the default
    /// Excel VBA names of `Sheet1`, `Sheet2`, etc.
    ///
    /// See also the
    /// [`Workbook::set_vba_name()`](crate::Workbook::set_vba_name()) method for
    /// setting the workbook VBA name.
    ///
    /// The name must be a valid Excel VBA object name as defined by the
    /// following rules:
    ///
    /// - The name must be less than 32 characters.
    /// - The name can only contain word characters: letters, numbers and
    ///   underscores.
    /// - The name must start with a letter.
    /// - The name cannot be blank.
    ///
    /// The name must be also be unique across the worksheets in the workbook.
    ///
    /// # Parameters
    ///
    /// - `name`: The vba name. It must follow the Excel rules, shown above.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::VbaNameError`] - The name doesn't meet one of Excel's
    ///   criteria, shown above.
    ///
    pub fn set_vba_name(&mut self, name: impl Into<String>) -> Result<&mut Worksheet, XlsxError> {
        let name = name.into();
        utility::validate_vba_name(&name)?;
        self.vba_codename = Some(name);

        Ok(self)
    }

    /// Set the default string used for NaN values.
    ///
    /// Excel doesn't support storing `NaN` (Not a Number) values. If a `NAN` is
    /// generated as the result of a calculation Excel stores and displays the
    /// error `#NUM!`. However, this error isn't usually used outside of a
    /// formula result and it isn't stored as a number.
    ///
    /// In order to deal with [`f64::NAN`] in a reasonable way `rust_xlsxwriter`
    /// writes it as the string "NAN". The `set_nan_value()` method allows you
    /// to override this default value.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for NaN values.
    ///
    /// # Examples
    ///
    /// The following example demonstrates handling NaN and Infinity values and
    /// also setting custom string representations.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_set_nan_string.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write NaN and Infinity default values.
    ///     worksheet.write(0, 0, "Default:")?;
    ///     worksheet.write(0, 1, f64::NAN)?;
    ///     worksheet.write(1, 1, f64::INFINITY)?;
    ///     worksheet.write(2, 1, f64::NEG_INFINITY)?;
    ///
    ///     // Overwrite the default values.
    ///     worksheet.set_nan_value("Nan");
    ///     worksheet.set_infinity_value("Infinity");
    ///     worksheet.set_neg_infinity_value("NegInfinity");
    ///
    ///     // Write NaN and Infinity custom values.
    ///     worksheet.write(4, 0, "Custom:")?;
    ///     worksheet.write(4, 1, f64::NAN)?;
    ///     worksheet.write(5, 1, f64::INFINITY)?;
    ///     worksheet.write(6, 1, f64::NEG_INFINITY)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_nan_string.png">
    ///
    pub fn set_nan_value(&mut self, value: impl Into<String>) -> &mut Worksheet {
        self.nan = value.into();
        self
    }

    /// Set the default string used for Infinite values.
    ///
    /// Excel doesn't support storing `Inf` (infinity) values. If an `Inf` is
    /// generated as the result of a calculation Excel stores and displays the
    /// error `#DIV/0`. However, this error isn't usually used outside of a
    /// formula result and it isn't stored as a number.
    ///
    /// In order to deal with [`f64::INFINITY`] in a reasonable way
    /// `rust_xlsxwriter` writes it as the string "INF". The
    /// `set_infinite_value()` method allows you to override this default value.
    ///
    /// See the example for [`Worksheet::set_nan_value()`] above.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for `Inf` values.
    ///
    pub fn set_infinity_value(&mut self, value: impl Into<String>) -> &mut Worksheet {
        self.infinity = value.into();
        self
    }

    /// Set the default string used for negative Infinite values.
    ///
    /// Excel doesn't support storing `-Inf` (negative infinity) values. If a
    /// `-Inf` is generated as the result of a calculation Excel stores and
    /// displays the error `#DIV/0`. However, this error isn't usually used
    /// outside of a formula result and it isn't stored as a number.
    ///
    ///
    /// In order to deal with [`f64::NEG_INFINITY`] in a reasonable way
    /// `rust_xlsxwriter` writes it as the string "-INF". The
    /// `set_infinite_value()` method method allows you to override this default
    /// value.
    ///
    /// See the example for [`Worksheet::set_nan_value()`] above.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for `-Inf` values.
    ///
    pub fn set_neg_infinity_value(&mut self, value: impl Into<String>) -> &mut Worksheet {
        self.neg_infinity = value.into();
        self
    }

    /// Ignore an Excel error or warning in a worksheet cell.
    ///
    /// Excel flags a number of data errors and inconsistencies with a a small
    /// green triangle in the top left hand corner of the cell. For example the
    /// following causes a warning of "Number Stored as Text":
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_ignore_error1.rs
    /// #
    /// # use rust_xlsxwriter::{Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write a string that looks like a number.
    ///     worksheet.write_string(1, 2, "123")?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_ignore_error1.png">
    ///
    /// These warnings can be useful indicators that there is an issue in the
    /// spreadsheet but sometimes it is preferable to turn them off. At the file
    /// level these errors can be ignored for a cell or cell range using
    /// `Worksheet::ignore_error()` and [`Worksheet::ignore_error_range()`] (see
    /// below).
    ///
    /// The errors and warnings that can be turned off at the file level are
    /// represented by the [`IgnoreError`] enum values. These equate, with some
    /// minor exceptions, to the error categories shown in the Excel Error
    /// Checking dialog:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/ignore_errors_dialog.png">
    ///
    /// (Note: some of the items shown in the above dialog such as "Misleading
    /// Number Formats" aren't saved in the output file by Excel and can't be
    /// turned off permanently.)
    ///
    /// <br>
    ///
    /// The `Worksheet::ignore_error()` method can be called repeatedly to
    /// ignore errors in different cells but **Excel only allows one ignored
    /// error per cell**.
    ///
    /// An error can be turned off for a range of cells using the
    /// [`Worksheet::ignore_error_range()`] method (see below).
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `error_type`: An [`IgnoreError`] enum value.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::ParameterError`] - Parameter error if more than one rule
    ///   is added to the same cell.
    ///
    /// # Examples
    ///
    /// An example of turning off worksheet cells errors/warnings using using
    /// the `rust_xlsxwriter` library.
    ///
    /// ```
    /// # // This code is available in examples/app_ignore_errors.rs
    /// #
    /// # use rust_xlsxwriter::{Format, IgnoreError, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a format to use in descriptions.
    /// #     let bold = Format::new().set_bold();
    /// #
    /// #     // Make the column wider for clarity.
    /// #     worksheet.set_column_width(1, 16)?;
    /// #
    /// #     // Write some descriptions for the cells.
    /// #     worksheet.write_with_format(1, 1, "Warning:", &bold)?;
    /// #     worksheet.write_with_format(2, 1, "Warning turned off:", &bold)?;
    /// #     worksheet.write_with_format(4, 1, "Warning:", &bold)?;
    /// #     worksheet.write_with_format(5, 1, "Warning turned off:", &bold)?;
    /// #
    ///     // Write strings that looks like numbers. This will cause an Excel warning.
    ///     worksheet.write_string(1, 2, "123")?;
    ///     worksheet.write_string(2, 2, "123")?;
    ///
    ///     // Write a divide by zero formula. This will also cause an Excel warning.
    ///     worksheet.write_formula(4, 2, "=1/0")?;
    ///     worksheet.write_formula(5, 2, "=1/0")?;
    ///
    ///     // Turn off some of the warnings:
    ///     worksheet.ignore_error(2, 2, IgnoreError::NumberStoredAsText)?;
    ///     worksheet.ignore_error(5, 2, IgnoreError::FormulaError)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("ignore_errors.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/app_ignore_errors.png">
    ///
    pub fn ignore_error(
        &mut self,
        row: RowNum,
        col: ColNum,
        error_type: IgnoreError,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.ignore_error_range(row, col, row, col, error_type)
    }

    /// Ignore an Excel error or warning in a range of worksheet cells.
    ///
    /// See [`Worksheet::ignore_error()`] above for a detailed explanation of
    /// Excel worksheet errors.
    ///
    /// The `Worksheet::ignore_error_range()` method can be used to ignore an
    /// error in a range, a row, a column, or the entire worksheet and it can be
    /// called repeatedly to ignore errors in different cells ranges. It should
    /// be noted however that **Excel only allows one ignored error per cell**.
    /// The `rust_xlsxwriter` library verifies that multiple rules aren't added
    /// to the same cell or cell range and will raise an error if this occurs.
    /// However it doesn't currently verify whether cells within ranges overlap.
    /// It is up to the user to ensure that this doesn't happen when using
    /// ranges.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `error_type`: An [`IgnoreError`] enum value.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row or column is larger
    ///   than the last row or column.
    /// - [`XlsxError::ParameterError`] - Parameter error if more than one rule
    ///   is added to the same range.
    ///
    pub fn ignore_error_range(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        error_type: IgnoreError,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions_only(first_row, first_col)
            || !self.check_dimensions_only(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        let range = utility::cell_range(first_row, first_col, last_row, last_col);

        // Check if the range was used previously. Note, that this doesn't
        // current check for range overlaps. One of the rules in an overlap will
        // be ignored since Excel only allows one rule per cell.
        if !self
            .ignored_error_cells
            .insert((first_row, first_col, last_row, last_col))
        {
            return Err(XlsxError::ParameterError(format!(
                "Excel only allows on rule per cell. Cell/Range '{range}' was previously used."
            )));
        }

        // Store the range or append it to an existing range.
        self.ignored_errors
            .entry(error_type)
            .and_modify(|sqref| *sqref = format!("{sqref} {range}"))
            .or_insert(range);

        Ok(self)
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    // Set the write ahead mode when using constant memory.
    fn set_writing_ahead(&mut self, enable: bool) {
        if self.use_constant_memory {
            self.is_writing_ahead = enable;
        }
    }

    // Get the minimum row number for the dimension check/set.
    fn get_min_row(&self) -> RowNum {
        if self.dimensions.first_row == ROW_MAX {
            0
        } else {
            self.dimensions.first_row
        }
    }

    // Get the minimum col number for the dimension check/set.
    fn get_min_col(&self) -> ColNum {
        if self.dimensions.first_col == COL_MAX {
            0
        } else {
            self.dimensions.first_col
        }
    }

    // Hide any rows in the autofilter range that don't match the autofilter
    // conditions, like Excel does at runtime.
    pub(crate) fn hide_autofilter_rows(&mut self) {
        if self.filter_conditions.is_empty() || self.filter_automatic_off {
            return;
        }

        // Get the range that the autofilter applies to.
        let first_row = self.autofilter_defined_name.first_row + 1; // Skip header.
        let last_row = self.autofilter_defined_name.last_row;

        for col_num in self.filter_conditions.clone().keys() {
            // Iterate through each column filter conditions.
            let filter_condition = self.filter_conditions.get(col_num).unwrap().clone();
            for row_num in first_row..=last_row {
                if filter_condition.is_list_filter {
                    // Handle list filters.
                    if !self.row_matches_list_filter(row_num, *col_num, &filter_condition) {
                        self.set_row_hidden(row_num).unwrap();
                    }
                } else {
                    // Handle custom filters.
                    if !self.row_matches_custom_filters(row_num, *col_num, &filter_condition) {
                        self.set_row_hidden(row_num).unwrap();
                    }
                }
            }
        }
    }

    // Check if the data in a cell matches one of the values in the list of
    // filter conditions (which in the list filter case is a list of strings or
    // number values).
    //
    // Excel trims leading and trailing space and then does a lowercase
    // comparison. It also matches numbers against "numbers stored as strings".
    // It also treats "blanks" as empty cells but also any string that is
    // composed of whitespace. See the test cases for examples. We try to match
    // all these conditions.
    fn row_matches_list_filter(
        &self,
        row_num: RowNum,
        col_num: ColNum,
        filter_condition: &FilterCondition,
    ) -> bool {
        let mut has_cell_data = false;

        if let Some(columns) = self.data_table.get(&row_num) {
            if let Some(cell) = columns.get(&col_num) {
                has_cell_data = true;

                match cell {
                    CellType::String { string, .. }
                    | CellType::RichString {
                        raw_string: string, ..
                    } => {
                        let cell_string = string.clone().to_lowercase().trim().to_string();

                        for filter in &filter_condition.list {
                            if cell_string == filter.string.to_lowercase().trim() {
                                return true;
                            }
                        }

                        if filter_condition.should_match_blanks && cell_string.is_empty() {
                            return true;
                        }
                    }
                    CellType::Number { number, .. } => {
                        for filter in &filter_condition.list {
                            if filter.data_type == FilterDataType::Number
                                && number == &filter.number
                            {
                                return true;
                            }
                        }
                    }
                    CellType::Blank { .. } => {
                        if filter_condition.should_match_blanks {
                            return true;
                        }
                    }
                    // We don't currently try to handle matching any other data types.
                    _ => {}
                };
            }
        }

        // If there is no cell data then that qualifies as Blanks in Excel.
        if !has_cell_data && filter_condition.should_match_blanks {
            return true;
        }

        // If none of the conditions match then we return false and hide the row.
        false
    }

    // Check if the data in a cell matches one of the conditions and values is a
    // custom filter. Excel allows 1 or 2 custom filters. We check for each
    // filter and evaluate the result(s) with the user defined and/or condition.
    fn row_matches_custom_filters(
        &self,
        row_num: RowNum,
        col_num: ColNum,
        filter_condition: &FilterCondition,
    ) -> bool {
        let condition1;
        let condition2;

        if let Some(data) = &filter_condition.custom1 {
            condition1 = self.row_matches_custom_filter(row_num, col_num, data);
        } else {
            condition1 = false;
        }

        if let Some(data) = &filter_condition.custom2 {
            condition2 = self.row_matches_custom_filter(row_num, col_num, data);
        } else {
            return condition1;
        }

        if filter_condition.apply_logical_or {
            condition1 || condition2
        } else {
            condition1 && condition2
        }
    }

    // Check if the data in a cell matches one custom filter.
    //
    // Excel trims leading and trailing space and then does a lowercase
    // comparison. It also matches numbers against "numbers stored as strings".
    // It also applies the comparison operators to strings. However, it doesn't
    // apply the string criteria (like contains()) to numbers (unless they are
    // stored as strings).
    fn row_matches_custom_filter(
        &self,
        row_num: RowNum,
        col_num: ColNum,
        filter: &FilterData,
    ) -> bool {
        if let Some(columns) = self.data_table.get(&row_num) {
            if let Some(cell) = columns.get(&col_num) {
                match cell {
                    CellType::String { string, .. }
                    | CellType::RichString {
                        raw_string: string, ..
                    } => {
                        let cell_string = string.clone().to_lowercase().trim().to_string();
                        let filter_string = filter.string.to_lowercase().trim().to_string();

                        match filter.criteria {
                            FilterCriteria::EqualTo => return cell_string == filter_string,
                            FilterCriteria::NotEqualTo => return cell_string != filter_string,
                            FilterCriteria::LessThan => return cell_string < filter_string,
                            FilterCriteria::GreaterThan => return cell_string > filter_string,
                            FilterCriteria::LessThanOrEqualTo => {
                                return cell_string <= filter_string
                            }
                            FilterCriteria::GreaterThanOrEqualTo => {
                                return cell_string >= filter_string
                            }
                            FilterCriteria::EndsWith => {
                                return cell_string.ends_with(&filter_string)
                            }
                            FilterCriteria::DoesNotEndWith => {
                                return !cell_string.ends_with(&filter_string)
                            }
                            FilterCriteria::BeginsWith => {
                                return cell_string.starts_with(&filter_string)
                            }
                            FilterCriteria::DoesNotBeginWith => {
                                return !cell_string.starts_with(&filter_string)
                            }
                            FilterCriteria::Contains => {
                                return cell_string.contains(&filter_string)
                            }
                            FilterCriteria::DoesNotContain => {
                                return !cell_string.contains(&filter_string)
                            }
                        }
                    }
                    CellType::Number { number, .. } => {
                        if filter.data_type == FilterDataType::Number {
                            match filter.criteria {
                                FilterCriteria::EqualTo => return *number == filter.number,
                                FilterCriteria::LessThan => return *number < filter.number,
                                FilterCriteria::NotEqualTo => return *number != filter.number,
                                FilterCriteria::GreaterThan => return *number > filter.number,
                                FilterCriteria::LessThanOrEqualTo => {
                                    return *number <= filter.number
                                }
                                FilterCriteria::GreaterThanOrEqualTo => {
                                    return *number >= filter.number
                                }
                                _ => {}
                            }
                        }
                    }
                    CellType::Blank { .. } => {
                        // We need to handle "match non-blanks" as a special condition.
                        // Excel converts this to a custom filter of `!= " "`.
                        if filter.criteria == FilterCriteria::NotEqualTo && filter.string == " " {
                            return false;
                        }
                    }
                    _ => {
                        // Any existing non-blank cell should match the "non-blanks" criteria
                        // explained above.
                        if filter.criteria == FilterCriteria::NotEqualTo && filter.string == " " {
                            return true;
                        }
                    }
                };
            }
        }

        false
    }

    // Process pagebreaks to sort them, remove duplicates and check the number
    // is within the Excel limit.
    pub(crate) fn process_pagebreaks(breaks: &[u32]) -> Result<Vec<u32>, XlsxError> {
        let unique_breaks: HashSet<u32> = breaks.iter().copied().collect();
        let mut breaks: Vec<u32> = unique_breaks.into_iter().collect();
        breaks.sort_unstable();

        // Remove invalid 0 row/col.
        if breaks[0] == 0 {
            breaks.remove(0);
        }

        // The Excel 2007 specification says that the maximum number of page
        // breaks is 1026. However, in practice it is actually 1023.
        if breaks.len() > 1023 {
            let error =
                "Maximum number of horizontal or vertical pagebreaks allowed by Excel is 1023"
                    .to_string();
            return Err(XlsxError::ParameterError(error));
        }

        Ok(breaks)
    }

    // Store a number cell in the worksheet data table structure.
    fn store_number(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: impl Into<f64>,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.store_number_type(row, col, number.into(), format, false)
    }

    // Store a datetime cell in the worksheet data table structure.
    fn store_datetime(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: f64,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        self.store_number_type(row, col, number, format, true)
    }

    // Store a number/datetime cell in the worksheet data table structure.
    fn store_number_type(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: f64,
        format: Option<&Format>,
        is_datetime: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Excel doesn't have a NAN type/value so write a string instead.
        if number.is_nan() {
            return self.store_string(row, col, self.nan.clone(), format);
        }

        // Excel doesn't have an Infinity type/value so write a string instead.
        if number.is_infinite() {
            if number == f64::INFINITY {
                return self.store_string(row, col, self.infinity.clone(), format);
            }

            return self.store_string(row, col, self.neg_infinity.clone(), format);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => self.row_col_format(row, col),
        };

        // Create the appropriate cell type to hold the data.
        let cell = if is_datetime {
            CellType::DateTime { number, xf_index }
        } else {
            CellType::Number { number, xf_index }
        };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store a string cell in the worksheet data table structure.
    fn store_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: String,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Empty strings are ignored by Excel unless they have a format in which
        // case they are treated as a blank cell.
        if string.is_empty() {
            match format {
                Some(format) => return self.write_blank(row, col, format),
                None => return Ok(self),
            };
        }

        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        //  Check that the string is < Excel limit of 32767 chars.
        if string.chars().count() > MAX_STRING_LEN {
            return Err(XlsxError::MaxStringLengthExceeded);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => self.row_col_format(row, col),
        };

        // Create the appropriate cell type to hold the data.
        let cell = if self.use_inline_strings {
            CellType::InlineString {
                string: Arc::from(string),
                xf_index,
            }
        } else {
            let mut string_id = None;
            let string = Arc::from(string);

            if self.has_workbook_global_sst {
                let mut string_table = self.string_table.lock().unwrap();
                string_id = Some(string_table.shared_string_index(Arc::clone(&string)));
            }

            CellType::String {
                string,
                xf_index,
                string_id,
            }
        };

        if !self.use_inline_strings {
            self.has_local_string_table = true;
        }

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store a rich string cell in the worksheet data table structure.
    fn store_rich_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        raw_string: &str,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Empty strings are ignored by Excel unless they have a format in which
        // case they are treated as a blank cell.
        if string.is_empty() {
            match format {
                Some(format) => return self.write_blank(row, col, format),
                None => return Ok(self),
            };
        }

        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        //  Check that the string is < Excel limit of 32767 chars.
        if string.chars().count() > MAX_STRING_LEN {
            return Err(XlsxError::MaxStringLengthExceeded);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => self.row_col_format(row, col),
        };

        // Create the appropriate cell type to hold the data.
        let mut string_id = None;
        let string = Arc::from(string);
        let raw_string = Arc::from(raw_string);

        if self.has_workbook_global_sst {
            let mut string_table = self.string_table.lock().unwrap();
            string_id = Some(string_table.shared_string_index(Arc::clone(&string)));
        }

        let cell = CellType::RichString {
            string,
            xf_index,
            raw_string,
            string_id,
        };

        self.insert_cell(row, col, cell);

        if !self.use_inline_strings {
            self.has_local_string_table = true;
        }

        Ok(self)
    }

    // Store a formula cell in the worksheet data table structure.
    fn store_formula(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: Formula,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Transfer to dynamic formula handling function.
        if formula.has_dynamic_function {
            return self.store_array_formula(row, col, row, col, formula, None, true);
        }

        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => self.row_col_format(row, col),
        };

        // Set the formula result to the default or user defined
        let result = if formula.result.is_empty() {
            self.default_result.clone()
        } else {
            formula.result.clone()
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Formula {
            formula: Box::from(formula.formula_string),
            xf_index,
            result,
        };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store an array formula cell in the worksheet data table structure.
    #[allow(clippy::too_many_arguments)]
    fn store_array_formula(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        formula: Formula,
        format: Option<&Format>,
        is_dynamic: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check rows and cols are in the allowed range.
        if !self.check_dimensions(first_row, first_col)
            || !self.check_dimensions(last_row, last_col)
        {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row || first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => self.row_col_format(first_row, first_col),
        };

        // Create the array range reference.
        let range = utility::cell_range(first_row, first_col, last_row, last_col);

        // Check for a dynamic function in a standard static array formula.
        let mut is_dynamic = is_dynamic;
        if !is_dynamic && formula.has_dynamic_function {
            is_dynamic = true;
        }

        if is_dynamic {
            self.has_dynamic_arrays = true;
        }

        // Set the formula result to the default or user defined
        let result = if formula.result.is_empty() {
            self.default_result.clone()
        } else {
            formula.result.clone()
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::ArrayFormula {
            formula: Box::from(formula.formula_string),
            xf_index,
            result,
            is_dynamic,
            range: range.into_boxed_str(),
        };

        self.insert_cell(first_row, first_col, cell);

        // Pad out the rest of the range with 0 result cells. We split this into
        // the first row and subsequent rows to allow us to handle "constant
        // mode" write-ahead.
        for col in first_col + 1..=last_col {
            match format {
                Some(format) => self.write_number_with_format(first_row, col, 0, format)?,
                None => self.write_number(first_row, col, 0)?,
            };
        }
        self.set_writing_ahead(true);
        for row in first_row + 1..=last_row {
            for col in first_col..=last_col {
                match format {
                    Some(format) => self.write_number_with_format(row, col, 0, format)?,
                    None => self.write_number(row, col, 0)?,
                };
            }
        }
        self.set_writing_ahead(false);

        Ok(self)
    }

    // Store a blank cell in the worksheet data table structure.
    fn store_blank(
        &mut self,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object.
        let mut xf_index = self.format_xf_index(format);

        // If the cell has the default format we need to check if it should have
        // a row/col intersection format.
        if xf_index == 0 {
            xf_index = self.row_col_format(row, col);
        }

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Blank { xf_index };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store a boolean cell in the worksheet data table structure.
    fn store_boolean(
        &mut self,
        row: RowNum,
        col: ColNum,
        boolean: bool,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => self.row_col_format(row, col),
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Boolean { boolean, xf_index };

        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // Store a url and associated properties. Urls in Excel are stored in a
    // number of places: they are written as a string similar to
    // write_string_with_format(), they are written in the <hyperlinks> element
    // within the worksheet, and they are referenced in the worksheet.rels file.
    fn store_url(
        &mut self,
        row: RowNum,
        col: ColNum,
        url: &Url,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        let mut hyperlink = url.clone();
        hyperlink.initialize()?;

        match format {
            Some(format) => {
                self.write_string_with_format(row, col, &hyperlink.user_text, format)?
            }
            None => {
                let hyperlink_format = Format::new().set_hyperlink();
                self.write_string_with_format(row, col, &hyperlink.user_text, &hyperlink_format)?
            }
        };

        self.hyperlinks.insert((row, col), hyperlink);

        Ok(self)
    }

    // Store a reference to an embedded cell image.
    fn store_embedded_image(
        &mut self,
        row: RowNum,
        col: ColNum,
        image: &Image,
        format: Option<&Format>,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Since embedded images need to have an associated cell error type they
        // can only be written in the current row for constant memory mode.
        if self.use_constant_memory && row < self.current_row {
            return Err(XlsxError::ParameterError(format!(
                "Cannot embed image to previously written row {row} in 'constant memory' mode."
            )));
        }

        let image_id = match self.embedded_image_ids.get(&image.hash) {
            Some(image_id) => *image_id,
            None => {
                let image_id = self.embedded_image_ids.len() as u32;
                self.embedded_image_ids.insert(image.hash.clone(), image_id);
                self.embedded_images.push(image.clone());
                image_id
            }
        };

        // Check for alt text in the image.
        if !image.alt_text.is_empty() {
            self.has_embedded_image_descriptions = true;
        }

        // Store the used image type for the Content Type file.
        self.image_types[image.image_type.clone() as usize] = true;

        // Store the image hyperlink, if any.
        if let Some(url) = &image.url {
            let mut hyperlink = url.clone();
            hyperlink.rel_display = true;

            self.hyperlinks.insert((row, col), hyperlink);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => self.format_xf_index(format),
            None => match image.url {
                Some(_) => {
                    let format = Format::new().set_hyperlink();
                    self.format_xf_index(&format)
                }
                None => self.row_col_format(row, col),
            },
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Error {
            xf_index,
            value: image_id,
        };

        // Store the cell error value.
        self.insert_cell(row, col, cell);

        Ok(self)
    }

    // A rich string is handled in Excel like any other shared string except
    // that it has inline font markup within the string. To generate the
    // required font xml we use an instance of the Style struct.
    #[allow(clippy::similar_names)]
    fn get_rich_string(segments: &[(&Format, &str)]) -> Result<(String, String), XlsxError> {
        // Check that there is at least one segment tuple.
        if segments.is_empty() {
            let error = "Rich string must contain at least 1 (&Format, &str) tuple.";
            return Err(XlsxError::ParameterError(error.to_string()));
        }

        // Create a Style struct object to generate the font xml.
        let xf_formats: Vec<Format> = vec![];
        let dxf_formats: Vec<Format> = vec![];
        let mut styler = Styles::new(
            &xf_formats,
            &dxf_formats,
            0,
            0,
            0,
            vec![],
            false,
            false,
            true,
        );
        let mut raw_string = String::new();

        let mut first_segment = true;
        for (format, string) in segments {
            // Excel doesn't allow empty string segments in a rich string.
            if string.is_empty() {
                let error = "Strings in rich string (&Format, &str) tuples cannot be blank.";
                return Err(XlsxError::ParameterError(error.to_string()));
            }

            // Accumulate the string segments into a unformatted string.
            raw_string.push_str(string);

            let whitespace = ['\t', '\n', ' '];
            let attributes = if string.starts_with(whitespace) || string.ends_with(whitespace) {
                vec![("xml:space", "preserve")]
            } else {
                vec![]
            };

            // First segment doesn't require a font run for the default format.
            if format.is_default() && first_segment {
                xml_start_tag_only(&mut styler.writer, "r");
                xml_data_element(&mut styler.writer, "t", string, &attributes);
                xml_end_tag(&mut styler.writer, "r");
            } else {
                xml_start_tag_only(&mut styler.writer, "r");
                styler.write_font(&format.font, false);
                xml_data_element(&mut styler.writer, "t", string, &attributes);
                xml_end_tag(&mut styler.writer, "r");
            }
            first_segment = false;
        }

        let got = xmlwriter::cursor_to_string(&styler.writer);

        Ok((got, raw_string))
    }

    // Insert a cell value into the worksheet data table structure.
    fn insert_cell(&mut self, row: RowNum, col: ColNum, cell: CellType) {
        if self.use_constant_memory {
            #[cfg(feature = "constant_memory")]
            {
                // Ignore already flushed/written rows.
                if row < self.current_row {
                    eprintln!(
                        "Ignoring write to previously written row {row} in 'constant memory' mode."
                    );
                    return;
                }

                // If this is a new row then either buffer the data when writing
                // ahead or flush the previous row.
                if row > self.current_row {
                    if self.is_writing_ahead {
                        // Store cell in the write-ahead buffer.
                        Self::insert_cell_to_table(row, col, cell, &mut self.write_ahead);
                        return;
                    }

                    self.flush_to_row(row);
                }

                // Store new constant memory data in the current row of the data table.
                Self::insert_cell_to_table(row, col, cell, &mut self.data_table);
            }
        } else {
            // In standard-memory mode all cell data is stored.
            Self::insert_cell_to_table(row, col, cell, &mut self.data_table);
        }
    }

    // Add a cell to one of the data tables.
    fn insert_cell_to_table(
        row: RowNum,
        col: ColNum,
        cell: CellType,
        data_table: &mut BTreeMap<RowNum, BTreeMap<ColNum, CellType>>,
    ) {
        match data_table.entry(row) {
            Entry::Occupied(mut entry) => {
                // The row already exists. Insert/replace column value.
                let columns = entry.get_mut();
                columns.insert(col, cell);
            }
            Entry::Vacant(entry) => {
                // The row doesn't exist, create a new row with columns and insert
                // the cell value.
                let columns = BTreeMap::from([(col, cell)]);
                entry.insert(columns);
            }
        }
    }

    // Insert a cell format value into the worksheet data table structure. This
    // function creates a new blank cell if no other cell value exists.
    fn insert_cell_format(&mut self, row: RowNum, col: ColNum, format_id: u32) {
        match self.data_table.entry(row) {
            Entry::Occupied(mut entry) => {
                // The row already exists.
                let columns = entry.get_mut();

                match columns.get_mut(&col) {
                    // The cell exists, so update the format.
                    Some(cell) => match cell {
                        CellType::Blank { xf_index, .. }
                        | CellType::Error { xf_index, .. }
                        | CellType::String { xf_index, .. }
                        | CellType::Number { xf_index, .. }
                        | CellType::Boolean { xf_index, .. }
                        | CellType::Formula { xf_index, .. }
                        | CellType::DateTime { xf_index, .. }
                        | CellType::RichString { xf_index, .. }
                        | CellType::InlineString { xf_index, .. }
                        | CellType::ArrayFormula { xf_index, .. } => {
                            *xf_index = format_id;
                        }
                    },

                    // The cell doesn't exist so add a blank cell with the format.
                    None => {
                        let cell = CellType::Blank {
                            xf_index: format_id,
                        };

                        columns.insert(col, cell);
                    }
                }
            }
            Entry::Vacant(entry) => {
                // The row doesn't exist, create a new row and insert a blank
                // cell if it has a non-default format.
                if format_id > 0 {
                    let cell = CellType::Blank {
                        xf_index: format_id,
                    };

                    let columns = BTreeMap::from([(col, cell)]);
                    entry.insert(columns);
                }
            }
        }
    }

    // Update the format index in a worksheet cell. This function ignores
    // non-existing cells (unlike the previous function).
    fn update_cell_format(&mut self, row: RowNum, col: ColNum, format_id: u32) {
        let Some(columns) = self.data_table.get_mut(&row) else {
            return;
        };

        let Some(cell) = columns.get_mut(&col) else {
            return;
        };

        match cell {
            CellType::Blank { xf_index, .. }
            | CellType::Error { xf_index, .. }
            | CellType::String { xf_index, .. }
            | CellType::Number { xf_index, .. }
            | CellType::Boolean { xf_index, .. }
            | CellType::Formula { xf_index, .. }
            | CellType::DateTime { xf_index, .. }
            | CellType::RichString { xf_index, .. }
            | CellType::InlineString { xf_index, .. }
            | CellType::ArrayFormula { xf_index, .. } => {
                *xf_index = format_id;
            }
        }
    }

    // Clear the data and formatting from a worksheet cell. Ignores non-existing
    // cells.
    fn clear_cell_internal(&mut self, row: RowNum, col: ColNum) {
        let Some(columns) = self.data_table.get_mut(&row) else {
            return;
        };

        columns.remove(&col);
    }

    // Clear the formatting from a worksheet cell. Ignores non-existing cells.
    fn clear_cell_format_internal(&mut self, row: RowNum, col: ColNum) {
        self.update_cell_format(row, col, 0);
    }

    // Set the border around a single cell.
    fn set_range_border_cell(
        &mut self,
        row: RowNum,
        col: ColNum,
        cell_format: &Format,
        border_format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let cell_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::All);

        self.set_cell_format(row, col, &cell_format)
    }

    // Set the border around a row of cells.
    fn set_range_border_row(
        &mut self,
        row: RowNum,
        first_col: ColNum,
        last_col: ColNum,
        cell_format: &Format,
        border_format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let left_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::RowLeft);

        let right_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::RowRight);

        let center_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::RowCenter);

        for col in first_col..=last_col {
            if col == first_col {
                self.set_cell_format(row, col, &left_format)?;
            } else if col == last_col {
                self.set_cell_format(row, col, &right_format)?;
            } else {
                self.set_cell_format(row, col, &center_format)?;
            }
        }

        Ok(self)
    }

    // Set the border around a column of cells.
    fn set_range_border_col(
        &mut self,
        first_row: RowNum,
        last_row: RowNum,
        col: ColNum,
        cell_format: &Format,
        border_format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let top_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::ColTop);

        let bottom_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::ColBottom);

        let center_format =
            Self::combined_border_format(cell_format, border_format, BorderPosition::ColCenter);

        for row in first_row..=last_row {
            if row == first_row {
                self.set_cell_format(row, col, &top_format)?;
            } else if row == last_row {
                self.set_cell_format(row, col, &bottom_format)?;
            } else {
                self.set_cell_format(row, col, &center_format)?;
            }
        }

        Ok(self)
    }

    // Set the border around a range of cells.
    fn set_range_border_range(
        &mut self,
        first_row: RowNum,
        last_row: RowNum,
        first_col: ColNum,
        last_col: ColNum,
        cell_format: &Format,
        border_format: &Format,
    ) -> Result<&mut Worksheet, XlsxError> {
        let top_left =
            Self::combined_border_format(cell_format, border_format, BorderPosition::TopLeft);

        let top_center =
            Self::combined_border_format(cell_format, border_format, BorderPosition::TopCenter);

        let top_right =
            Self::combined_border_format(cell_format, border_format, BorderPosition::TopRight);

        let center_left =
            Self::combined_border_format(cell_format, border_format, BorderPosition::CenterLeft);

        let center_center =
            Self::combined_border_format(cell_format, border_format, BorderPosition::CenterCenter);

        let center_right =
            Self::combined_border_format(cell_format, border_format, BorderPosition::CenterRight);

        let bottom_left =
            Self::combined_border_format(cell_format, border_format, BorderPosition::BottomLeft);

        let bottom_center =
            Self::combined_border_format(cell_format, border_format, BorderPosition::BottomCenter);

        let bottom_right =
            Self::combined_border_format(cell_format, border_format, BorderPosition::BottomRight);

        for row in first_row..=last_row {
            for col in first_col..=last_col {
                if row == first_row {
                    if col == first_col {
                        self.set_cell_format(row, col, &top_left)?;
                    } else if col == last_col {
                        self.set_cell_format(row, col, &top_right)?;
                    } else {
                        self.set_cell_format(row, col, &top_center)?;
                    }
                } else if row == last_row {
                    if col == first_col {
                        self.set_cell_format(row, col, &bottom_left)?;
                    } else if col == last_col {
                        self.set_cell_format(row, col, &bottom_right)?;
                    } else {
                        self.set_cell_format(row, col, &bottom_center)?;
                    }
                } else if col == first_col {
                    self.set_cell_format(row, col, &center_left)?;
                } else if col == last_col {
                    self.set_cell_format(row, col, &center_right)?;
                } else {
                    self.set_cell_format(row, col, &center_center)?;
                }
            }
        }

        Ok(self)
    }

    // Combine a cell format and a border format to create a new format that can
    // be used when creating a range border.
    fn combined_border_format(
        cell_format: &Format,
        border_format: &Format,
        position: BorderPosition,
    ) -> Format {
        let mut cell_format = cell_format.clone();

        // Top.
        match position {
            BorderPosition::All
            | BorderPosition::RowLeft
            | BorderPosition::RowCenter
            | BorderPosition::RowRight
            | BorderPosition::ColTop
            | BorderPosition::TopLeft
            | BorderPosition::TopCenter
            | BorderPosition::TopRight => {
                cell_format.borders.top_style = border_format.borders.top_style;
                cell_format.borders.top_color = border_format.borders.top_color;
            }
            _ => {}
        }

        // Bottom.
        match position {
            BorderPosition::All
            | BorderPosition::RowLeft
            | BorderPosition::RowCenter
            | BorderPosition::RowRight
            | BorderPosition::ColBottom
            | BorderPosition::BottomLeft
            | BorderPosition::BottomCenter
            | BorderPosition::BottomRight => {
                cell_format.borders.bottom_style = border_format.borders.bottom_style;
                cell_format.borders.bottom_color = border_format.borders.bottom_color;
            }
            _ => {}
        }

        // Left.
        match position {
            BorderPosition::All
            | BorderPosition::RowLeft
            | BorderPosition::ColTop
            | BorderPosition::ColCenter
            | BorderPosition::ColBottom
            | BorderPosition::TopLeft
            | BorderPosition::CenterLeft
            | BorderPosition::BottomLeft => {
                cell_format.borders.left_style = border_format.borders.left_style;
                cell_format.borders.left_color = border_format.borders.left_color;
            }
            _ => {}
        }

        // Right.
        match position {
            BorderPosition::All
            | BorderPosition::RowRight
            | BorderPosition::ColTop
            | BorderPosition::ColCenter
            | BorderPosition::ColBottom
            | BorderPosition::TopRight
            | BorderPosition::CenterRight
            | BorderPosition::BottomRight => {
                cell_format.borders.right_style = border_format.borders.right_style;
                cell_format.borders.right_color = border_format.borders.right_color;
            }
            _ => {}
        }

        cell_format
    }

    // Set the column width in character units. This also takes into account
    // whether the column width is auto-fitted or manually set.
    fn set_column_width_internal(
        &mut self,
        col: ColNum,
        width: impl Into<f64>,
        autofit: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        let width = width.into();

        // If the width is 0 then the Excel treats the column as hidden with
        // default width.
        if width == 0.0 {
            return self.set_column_hidden(col);
        }

        // Check if column is in the allowed range without updating dimensions.
        if col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Store the column width.
        self.store_column_width(col, width, autofit);

        Ok(self)
    }

    // Store the column width in Excel character units. Updates to the width can
    // come from the external user or from the internal autofit() routines.
    fn store_column_width(&mut self, col: ColNum, width: f64, autofit: bool) {
        // Excel has a maximum limit of 255 units for the column width.
        let mut width = width;
        if width > 255.0 {
            width = 255.0;
        }

        // Update an existing col metadata object or create a new one.
        match self.changed_cols.get_mut(&col) {
            Some(col_options) => {
                // Note, autofit() will only update a user defined value if is
                // greater than it. All other conditions are simple updates.
                if autofit && !col_options.autofit {
                    if width > col_options.width {
                        col_options.width = width;
                        col_options.autofit = true;
                    }
                } else {
                    col_options.width = width;
                    col_options.autofit = autofit;
                }
            }
            None => {
                // Create a new column metadata object.
                let col_options = ColOptions {
                    width,
                    xf_index: 0,
                    hidden: false,
                    level: 0,
                    collapsed: false,
                    autofit,
                    format: None,
                };
                self.changed_cols.insert(col, col_options);
            }
        }
    }

    // Check that row and col are within the allowed Excel range and store max
    // and min values for use in other methods/elements.
    fn check_dimensions(&mut self, row: RowNum, col: ColNum) -> bool {
        // Check that the row an column number are within Excel's ranges.
        if row >= ROW_MAX {
            return false;
        }
        if col >= COL_MAX {
            return false;
        }

        // Store any changes in worksheet dimensions.
        self.dimensions.first_row = cmp::min(self.dimensions.first_row, row);
        self.dimensions.first_col = cmp::min(self.dimensions.first_col, col);

        // In constant memory mode we also need to ensure the row is valid.
        if self.use_constant_memory {
            if row >= self.current_row {
                self.dimensions.last_row = cmp::max(self.dimensions.last_row, row);
                self.dimensions.last_col = cmp::max(self.dimensions.last_col, col);
            }
        } else {
            self.dimensions.last_row = cmp::max(self.dimensions.last_row, row);
            self.dimensions.last_col = cmp::max(self.dimensions.last_col, col);
        }

        true
    }

    // Check that row and col are within the allowed Excel range but don't
    // modify the worksheet cell range.
    #[allow(clippy::unused_self)]
    pub(crate) fn check_dimensions_only(&mut self, row: RowNum, col: ColNum) -> bool {
        // Check that the row an column number are within Excel's ranges.
        if row >= ROW_MAX {
            return false;
        }
        if col >= COL_MAX {
            return false;
        }

        true
    }

    // Lookup a Format to get a unique index that identifies it based on its
    // properties. This is either done from the global lookup table (if we have
    // a copy) or from a local lookup that we will reconcile with the parent
    // workbook later.
    fn format_xf_index(&mut self, format: &Format) -> u32 {
        if self.has_workbook_global_xfs {
            self.format_xf_index_global(format)
        } else {
            self.format_xf_index_local(format)
        }
    }

    // Look up the Format index from the workbook/global format index table via
    // the worksheet's Arc<RwLock<HashMap>> clone.
    fn format_xf_index_global(&mut self, format: &Format) -> u32 {
        // Try a read() lock first to see if the format is known.
        let xf_indices = self.workbook_xf_indices.read().expect("RwLock poisoned");

        if let Some(xf_index) = xf_indices.get(format) {
            return *xf_index;
        }

        // Index wasn't found, so drop the read() lock and get a write() lock to
        // add the format and create a new index.
        drop(xf_indices);

        // Add the new format and give it an index.
        let mut xf_indices = self.workbook_xf_indices.write().expect("RwLock poisoned");
        let xf_index = xf_indices.len() as u32;
        xf_indices.insert(format.clone(), xf_index);

        xf_index
    }

    // Look up the Format index from the worksheet/local format index table.
    // This will be reconciled with the parent workbook when the file is saved.
    fn format_xf_index_local(&mut self, format: &Format) -> u32 {
        match self.xf_indices.get_mut(format) {
            Some(xf_index) => *xf_index,
            None => {
                let xf_index = self.xf_formats.len() as u32;
                self.xf_formats.push(format.clone());
                self.xf_indices.insert(format.clone(), xf_index);
                if format.font.is_hyperlink {
                    self.has_hyperlink_style = true;
                }
                xf_index
            }
        }
    }

    // Get the remapped local to global format index.
    fn get_global_xf_index(&self, xf_index: u32) -> u32 {
        if self.has_workbook_global_xfs {
            xf_index
        } else {
            self.global_xf_indices[xf_index as usize]
        }
    }

    // This method emulates Excel's behavior when an unformatted cell occurs at
    // the intersection of a row and column format. Excel handles this case by
    // creating and applying a new format that is a combination of the row and
    // column formats.
    //
    // Cells that only have a row format or a column format are handled by the
    // get_cell_xf_index() method.
    //
    // The method is called for all unformatted cells so we try to exit early to
    // keep it as efficient as possible.
    //
    fn row_col_format(&mut self, row: RowNum, col: ColNum) -> u32 {
        // Check for both a row and column format and exit is one of them
        // doesn't exist.
        let Some(col_option) = self.changed_cols.get(&col) else {
            return 0;
        };
        let Some(col_format) = &col_option.format else {
            return 0;
        };
        let Some(row_option) = self.changed_rows.get(&row) else {
            return 0;
        };
        let Some(row_format) = &row_option.format else {
            return 0;
        };

        // Combine the row and column formats into a new format.
        let combined_format = Format::merge(row_format, col_format);

        // Return the index of the combined format.
        self.format_xf_index(&combined_format)
    }

    /// Get the local instance DXF id for a format.
    ///
    /// Get the local instance DXF id for a format. These indexes will be
    /// replaced by global/workbook indices before the worksheet is saved. DXF
    /// indexed are used for Tables and Conditional Formats.
    ///
    /// This method is public but hidden to allow test cases to mirror the
    /// creation order for DXF ids which is usually the reverse of the order of
    /// the XF instance ids.
    ///
    /// # Parameters
    ///
    /// `format` - The [`Format`] instance to register.
    ///
    #[doc(hidden)] // Set dxf_index (public for testing).
    pub fn format_dxf_index(&mut self, format: &Format) -> u32 {
        match self.dxf_indices.get_mut(format) {
            Some(dxf_index) => *dxf_index,
            None => {
                let dxf_index = self.dxf_formats.len() as u32;
                let mut dxf_format = format.clone();
                dxf_format.is_dxf_format = true;
                self.dxf_formats.push(dxf_format.clone());
                self.dxf_indices.insert(dxf_format, dxf_index);
                if format.font.is_hyperlink {
                    self.has_hyperlink_style = true;
                }
                dxf_index
            }
        }
    }

    // Set the mapping between the local format indices and the global/workbook
    // indices for cell formats.
    pub(crate) fn set_global_xf_indices(&mut self, workbook_xf_indices: &[u32]) {
        self.global_xf_indices = workbook_xf_indices.to_vec();
    }

    // Set the mapping between the local format indices and the global/workbook
    // indices for cell formats.
    pub(crate) fn set_global_dxf_indices(&mut self, workbook_dxf_indices: &[u32]) {
        self.global_dxf_indices = workbook_dxf_indices.to_vec();
    }

    // Translate the cell xf_index into a global/workbook format index. We also
    // need to make sure that an unformatted cell (xf_index == 0) takes the row
    // format (if it exists) or, failing that, the column format (if that
    // exists).
    fn get_cell_xf_index(
        &mut self,
        xf_index: u32,
        row_options: Option<&RowOptions>,
        col_num: ColNum,
    ) -> u32 {
        // The local cell format index.
        let mut xf_index = xf_index;

        // If it is zero the cell is unformatted and we check for a row format.
        if xf_index == 0 {
            if let Some(row_options) = row_options {
                xf_index = row_options.xf_index;
            }
        }

        // If it is still zero the row was unformatted so we check for a column
        // format.
        if xf_index == 0 {
            if let Some(col_options) = self.changed_cols.get(&col_num) {
                xf_index = col_options.xf_index;
            }
        }

        // Finally convert the local format index into a global/workbook index.
        if xf_index != 0 {
            xf_index = self.get_global_xf_index(xf_index);
        }

        xf_index
    }

    // Convert the image dimensions into drawing dimensions and add them to the
    // Drawing object. Also set the rel linkages between the files.
    pub(crate) fn prepare_worksheet_images(
        &mut self,
        image_ids: &mut HashMap<String, u32>,
        image_id: &mut u32,
        drawing_id: u32,
    ) {
        for (cell, image) in &self.images.clone() {
            let row = cell.0;
            let col = cell.1;
            let mut drawing_hyperlink = None;

            let image_id = match image_ids.get(&image.hash) {
                Some(image_id) => *image_id,
                None => {
                    *image_id += 1;
                    image_ids.insert(image.hash.clone(), *image_id);
                    *image_id
                }
            };

            // Handle optional hyperlink in the image.
            if let Some(hyperlink) = &image.url {
                let mut hyperlink = hyperlink.clone();

                let target = hyperlink.target();
                let target_mode = hyperlink.target_mode();

                let rel_id = match self.drawing_rel_ids.get(&hyperlink.url_link) {
                    Some(rel_id) => *rel_id,
                    None => {
                        let rel_id = 1 + self.drawing_rel_ids.len() as u32;
                        self.drawing_rel_ids
                            .insert(hyperlink.url_link.clone(), rel_id);

                        // Store the linkage to the drawings rels file.
                        self.drawing_relationships.push((
                            "hyperlink".to_string(),
                            target,
                            target_mode,
                        ));

                        rel_id
                    }
                };

                hyperlink.rel_id = rel_id;

                drawing_hyperlink = Some(hyperlink);
            }

            // Store the image references.
            let rel_id = match self.drawing_rel_ids.get(&image.hash) {
                Some(rel_id) => *rel_id,
                None => {
                    let rel_id = 1 + self.drawing_rel_ids.len() as u32;
                    self.drawing_rel_ids.insert(image.hash.clone(), rel_id);

                    // Store the linkage to the drawings rels file.
                    let image_name =
                        format!("../media/image{image_id}.{}", image.image_type.extension());
                    self.drawing_relationships.push((
                        "image".to_string(),
                        image_name,
                        String::new(),
                    ));

                    rel_id
                }
            };

            // Convert the image dimensions to drawing dimensions and store the
            // drawing object.
            let mut drawing_info = self.position_object_emus(row, col, image);
            drawing_info.rel_id = rel_id;
            drawing_info.url.clone_from(&drawing_hyperlink);
            self.drawing.drawings.push(drawing_info);

            // Store the used image type for the Content Type file.
            self.image_types[image.image_type.clone() as usize] = true;
        }

        // Store the linkage to the worksheets rels file.
        let drawing_name = format!("../drawings/drawing{drawing_id}.xml");
        self.drawing_object_relationships.push((
            "drawing".to_string(),
            drawing_name,
            String::new(),
        ));

        self.has_drawing_object_linkage = true;
    }

    // Set the relationship for the background image.
    pub(crate) fn prepare_background_image(&mut self, image_id: u32, image: &Image) {
        let image_name = format!("../media/image{image_id}.{}", image.image_type.extension());

        self.image_types[image.image_type.clone() as usize] = true;

        self.background_relationships
            .push(("image".to_string(), image_name, String::new()));
    }

    // Convert the shape dimensions into drawing dimensions and add them to
    // the Drawing object. Also set the rel linkages between the files.
    pub(crate) fn prepare_worksheet_shapes(&mut self, shape_id: u32, drawing_id: u32) {
        let mut shape_id = shape_id;

        for (cell, shape) in &self.shapes.clone() {
            let row = cell.0;
            let col = cell.1;
            let mut drawing_hyperlink = None;

            // Handle optional hyperlink in the shape.
            if let Some(hyperlink) = &shape.url {
                let mut hyperlink = hyperlink.clone();

                let target = hyperlink.target();
                let target_mode = hyperlink.target_mode();

                let rel_id = match self.drawing_rel_ids.get(&hyperlink.url_link) {
                    Some(rel_id) => *rel_id,
                    None => {
                        let rel_id = 1 + self.drawing_rel_ids.len() as u32;
                        self.drawing_rel_ids
                            .insert(hyperlink.url_link.clone(), rel_id);

                        // Store the linkage to the drawings rels file.
                        self.drawing_relationships.push((
                            "hyperlink".to_string(),
                            target,
                            target_mode,
                        ));

                        rel_id
                    }
                };

                hyperlink.rel_id = rel_id;

                drawing_hyperlink = Some(hyperlink);
            }

            // Convert the shape dimensions to drawing dimensions and store
            // the drawing object.
            let mut drawing_info = self.position_object_emus(row, col, shape);
            drawing_info.rel_id = shape_id;
            drawing_info.url.clone_from(&drawing_hyperlink);
            self.drawing.drawings.push(drawing_info);
            self.drawing.shapes.push(shape.clone());

            shape_id += 1;
        }

        // Store the linkage to the worksheets rels file.
        if self.drawing_object_relationships.is_empty() {
            let drawing_name = format!("../drawings/drawing{drawing_id}.xml");
            self.drawing_object_relationships.push((
                "drawing".to_string(),
                drawing_name,
                String::new(),
            ));

            self.has_drawing_object_linkage = true;
        }
    }

    // Set up images used in headers and footers. Excel handles these
    // differently from worksheet images and stores them in a VML file rather
    // than an Drawing file.
    pub(crate) fn prepare_header_footer_images(
        &mut self,
        image_ids: &mut HashMap<String, u32>,
        base_image_id: u32,
    ) {
        let mut rel_ids: HashMap<String, u32> = HashMap::new();
        for image in self.header_footer_images.clone().into_iter().flatten() {
            let image_id = match image_ids.get(&image.hash) {
                Some(image_id) => *image_id,
                None => {
                    let image_id = 1 + base_image_id + image_ids.len() as u32;
                    image_ids.insert(image.hash.clone(), image_id);
                    image_id
                }
            };

            let rel_id = match rel_ids.get(&image.hash) {
                Some(rel_id) => *rel_id,
                None => {
                    let rel_id = 1 + rel_ids.len() as u32;
                    rel_ids.insert(image.hash.clone(), rel_id);

                    // Store the linkage to the drawings rels file.
                    let image_name =
                        format!("../media/image{image_id}.{}", image.image_type.extension());
                    self.vml_drawing_relationships.push((
                        "image".to_string(),
                        image_name,
                        String::new(),
                    ));

                    rel_id
                }
            };

            // Convert the header image to a VmlInfo structure for storing in a
            // vmlDrawing file..
            let mut vml_info = image.vml_info();
            vml_info.rel_id = rel_id;

            // Store the header/footer vml data.
            self.header_footer_vml_info.push(vml_info);

            // Store the used image type for the Content Type file.
            self.image_types[image.image_type as usize] = true;
        }
    }

    // Store the vmlDrawingN.xml file linkage to the worksheets rels file.
    pub(crate) fn add_vml_drawing_rel_link(&mut self, drawing_id: u32) {
        let vml_drawing_name = format!("../drawings/vmlDrawing{drawing_id}.vml");
        self.drawing_object_relationships.push((
            "vmlDrawing".to_string(),
            vml_drawing_name,
            String::new(),
        ));
    }

    // Convert buttons into VML objects.
    pub(crate) fn prepare_vml_objects(&mut self, vml_data_id: u32, vml_shape_id: u32) -> u32 {
        let mut button_id = 1;
        let mut note_count = 0;

        // Modify Note visibility and author according to worksheet settings.
        for columns in self.notes.values_mut() {
            for note in columns.values_mut() {
                // Set all notes visible if required.
                if self.show_all_notes && note.is_visible.is_none() {
                    note.is_visible = Some(true);
                }

                // Check for a user defined author name.
                let Some(note_author) = &note.author else {
                    continue;
                };

                // Convert the name to an id.
                match self.note_authors.get(note_author) {
                    Some(id) => {
                        note.author_id = *id;
                    }
                    None => {
                        let id = self.note_authors.len();
                        self.note_authors.insert(note_author.clone(), id);
                        note.author_id = id;
                    }
                }
            }
        }

        // Convert the Note objects to VmlInfo objects, along with their dimensions.
        for (cell_row, columns) in &self.notes.clone() {
            for (cell_col, note) in columns {
                let note_row = note.row();
                let note_col = note.col();

                let mut vml_info = note.vml_info();
                vml_info.drawing_info = self.position_object_pixels(note_row, note_col, note);
                vml_info.row = *cell_row;
                vml_info.col = *cell_col;

                // Store the note vml data.
                self.comments_vml_info.push(vml_info);

                note_count += 1;
            }
        }

        // Convert the Button objects to VmlInfo objects, along with their dimensions.
        for ((row, col, _, _), button) in self.buttons.clone() {
            let mut button = button.clone();
            if button.name.is_empty() {
                button.name = format!("Button {button_id}");
            }

            if button.macro_name.is_empty() {
                button.macro_name = format!("[0]!Button{button_id}_Click");
            } else {
                button.macro_name = format!("[0]!{}", button.macro_name);
            }

            let mut vml_info = button.vml_info();
            vml_info.drawing_info = self.position_object_pixels(row, col, &button);

            // Store the button vml data.
            self.buttons_vml_info.push(vml_info);

            button_id += 1;
        }

        // The VML o:idmap data id contains a comma separated range when there
        // is more than one 1024 block of comments, like this: data="1,2".
        let mut oid_map = vml_data_id.to_string();

        for i in 0..note_count / 1024 {
            let next_id = vml_data_id + i + 1;
            oid_map = format!("{oid_map},{next_id}");
        }

        self.vml_data_id = oid_map;
        self.vml_shape_id = vml_shape_id;

        note_count
    }

    // Store the commentN.xml file linkage to the worksheets rels file.
    pub(crate) fn add_comment_rel_link(&mut self, comment_id: u32) {
        let comment_name = format!("../comments{comment_id}.xml");
        self.comment_relationships
            .push(("comments".to_string(), comment_name, String::new()));
    }

    // Convert the chart dimensions into drawing dimensions and add them to the
    // Drawing object. Also set the rel linkages between the files.
    pub(crate) fn prepare_worksheet_charts(&mut self, chart_id: u32, drawing_id: u32) {
        let mut chart_id = chart_id;
        for chart in self.charts.values_mut() {
            chart.id = chart_id;
            chart.add_axis_ids(chart_id);
            chart_id += 1;
        }

        let mut rel_id = self.drawing_relationships.len() as u32;
        self.drawing_rel_ids.insert("Chart".to_string(), rel_id);

        for (cell, chart) in &mut self.charts.clone() {
            let row = cell.0;
            let col = cell.1;

            let chart_id = chart.id;

            // Store the linkage to the charts rels file.
            let chart_name = format!("../charts/chart{chart_id}.xml");
            self.drawing_relationships
                .push(("chart".to_string(), chart_name, String::new()));

            // Convert the chart dimensions to drawing dimensions and store the
            // drawing object.
            let mut drawing_info = self.position_object_emus(row, col, chart);
            rel_id += 1;
            drawing_info.rel_id = rel_id;

            if self.is_chartsheet {
                drawing_info.drawing_type = DrawingType::ChartSheet;
                drawing_info.is_portrait = self.is_portrait;
            }

            self.drawing.drawings.push(drawing_info);
        }

        // Set the chartsheet chart protection.
        if self.is_chartsheet && self.protection_on {
            if let Some(chart) = self.charts.get_mut(&(0, 0, 0, 0)) {
                chart.protection_on = true;
            }
        }

        // Store the linkage to the worksheets rels file, if it hasn't already
        // been set by the one of the image preparation functions.
        if !self.has_drawing_object_linkage {
            let drawing_name = format!("../drawings/drawing{drawing_id}.xml");
            self.drawing_object_relationships.push((
                "drawing".to_string(),
                drawing_name,
                String::new(),
            ));
        }
    }

    // Set a unique table id for each table and also set the rel linkages
    // between the worksheet and table xml files.
    pub(crate) fn prepare_worksheet_tables(&mut self, mut table_id: u32) -> u32 {
        for table in &mut self.tables {
            table.index = table_id;
            if table.name.is_empty() {
                table.name = format!("Table{table_id}");
            }

            // Change any column format local index from local to global.
            for column in &mut table.columns {
                if let Some(format) = column.format.as_mut() {
                    format.dxf_index = self.global_dxf_indices[format.dxf_index as usize];
                }
            }

            self.table_relationships.push((
                "table".to_string(),
                format!("../tables/table{table_id}.xml"),
                String::new(),
            ));

            table_id += 1;
        }

        table_id
    }

    // Calculate the vertices that define the position of a graphical object
    // within the worksheet in EMUs. The vertices are expressed as English
    // Metric Units (EMUs). There are 12,700 EMUs per point. Therefore, 12,700 *
    // 3 /4 = 9,525 EMUs per pixel.
    fn position_object_emus<T>(&mut self, row: RowNum, col: ColNum, object: &T) -> DrawingInfo
    where
        T: DrawingObject,
    {
        let mut drawing_info = self.position_object_pixels(row, col, object);

        // Convert the pixel values to EMUs.
        drawing_info.to.col_offset = round_to_emus(drawing_info.to.col_offset);
        drawing_info.to.row_offset = round_to_emus(drawing_info.to.row_offset);

        drawing_info.from.col_offset = round_to_emus(drawing_info.from.col_offset);
        drawing_info.from.row_offset = round_to_emus(drawing_info.from.row_offset);

        drawing_info.col_absolute *= 9525;
        drawing_info.row_absolute *= 9525;

        drawing_info.width = round_to_emus(drawing_info.width);
        drawing_info.height = round_to_emus(drawing_info.height);

        drawing_info
    }

    // Calculate the vertices that define the position of a graphical object
    // within the worksheet in pixels.
    //
    //         +------------+------------+
    //         |     A      |      B     |
    //   +-----+------------+------------+
    //   |     |(x1,y1)     |            |
    //   |  1  |(A1)._______|______      |
    //   |     |    |              |     |
    //   |     |    |              |     |
    //   +-----+----|    OBJECT    |-----+
    //   |     |    |              |     |
    //   |  2  |    |______________.     |
    //   |     |            |        (B2)|
    //   |     |            |     (x2,y2)|
    //   +---- +------------+------------+
    //
    // Example of an object that covers some of the area from cell A1 to  B2.
    //
    // Based on the width and height of the object we need to calculate 8 values:
    //
    //     col_start, row_start, col_end, row_end, x1, y1, x2, y2.
    //
    // We also calculate the absolute x and y position of the top left vertex of
    // the object. This is required for images.
    //
    // The width and height of the cells that the object occupies can be
    // variable and have to be taken into account.
    //
    // The values of col_start and row_start are passed in from the calling
    // function. The values of col_end and row_end are calculated by subtracting
    // the width and height of the object from the width and height of the
    // underlying cells.
    //
    fn position_object_pixels<T>(&mut self, row: RowNum, col: ColNum, object: &T) -> DrawingInfo
    where
        T: DrawingObject,
    {
        let mut row_start: RowNum = row; // Row containing top left corner.
        let mut col_start: ColNum = col; // Column containing upper left corner.

        let mut x1: u32 = object.x_offset(); // Distance to left side of object.
        let mut y1: u32 = object.y_offset(); // Distance to top of object.

        let mut row_end: RowNum; // Row containing bottom right corner.
        let mut col_end: ColNum; // Column containing lower right corner.

        let mut x2: f64; // Distance to right side of object.
        let mut y2: f64; // Distance to bottom of object.

        let width = object.width_scaled(); // Width of object frame.
        let height = object.height_scaled(); // Height of object frame.

        let mut x_abs: u32 = 0; // Absolute distance to left side of object.
        let mut y_abs: u32 = 0; // Absolute distance to top  side of object.

        // Calculate the absolute x offset of the top-left vertex.
        for col in 0..col_start {
            x_abs += self.column_pixel_width(col, object.object_movement());
        }
        x_abs += x1;

        // Calculate the absolute y offset of the top-left vertex.
        for row in 0..row_start {
            y_abs += self.row_pixel_height(row, object.object_movement());
        }
        y_abs += y1;

        // Adjust start col for offsets that are greater than the col width.
        loop {
            let col_size = self.column_pixel_width(col_start, object.object_movement());
            if x1 >= col_size {
                x1 -= col_size;
                col_start += 1;
            } else {
                break;
            }
        }

        // Adjust start row for offsets that are greater than the row height.
        loop {
            let row_size = self.row_pixel_height(row_start, object.object_movement());
            if y1 >= row_size {
                y1 -= row_size;
                row_start += 1;
            } else {
                break;
            }
        }

        // Initialize end cell to the same as the start cell.
        col_end = col_start;
        row_end = row_start;

        // Calculate the end vertices.
        x2 = width + f64::from(x1);
        y2 = height + f64::from(y1);

        // Subtract the underlying cell widths to find the end cell.
        loop {
            let col_size = f64::from(self.column_pixel_width(col_end, object.object_movement()));
            if x2 >= col_size {
                x2 -= col_size;
                col_end += 1;
            } else {
                break;
            }
        }

        //Subtract the underlying cell heights to find the end cell.
        loop {
            let row_size = f64::from(self.row_pixel_height(row_end, object.object_movement()));
            if y2 >= row_size {
                y2 -= row_size;
                row_end += 1;
            } else {
                break;
            }
        }

        // Create structs to hold the drawing information.
        let from = DrawingCoordinates {
            col: u32::from(col_start),
            row: row_start,
            col_offset: f64::from(x1),
            row_offset: f64::from(y1),
        };

        let to = DrawingCoordinates {
            col: u32::from(col_end),
            row: row_end,
            col_offset: x2,
            row_offset: y2,
        };

        DrawingInfo {
            from,
            to,
            col_absolute: u64::from(x_abs),
            row_absolute: u64::from(y_abs),
            width,
            height,
            name: object.name(),
            description: object.alt_text(),
            decorative: object.decorative(),
            object_movement: object.object_movement(),
            drawing_type: object.drawing_type(),
            rel_id: 0,
            url: None,
            is_portrait: false,
        }
    }

    // Convert the width of a cell from character units to pixels. Excel rounds
    // the column width to the nearest pixel.
    fn column_pixel_width(&mut self, col: ColNum, position: ObjectMovement) -> u32 {
        let max_digit_width = 7.0_f64;
        let padding = 5.0_f64;

        match self.changed_cols.get(&col) {
            Some(col_options) => {
                let pixel_width = col_options.width;
                let hidden = col_options.hidden;

                if hidden && position != ObjectMovement::MoveAndSizeWithCellsAfter {
                    // A hidden column is treated as having a width of zero unless
                    // the "object_movement" is MoveAndSizeWithCellsAfter.
                    0u32
                } else if pixel_width < 1.0 {
                    (pixel_width * (max_digit_width + padding)).round() as u32
                } else {
                    (pixel_width * max_digit_width).round() as u32 + padding as u32
                }
            }
            // If the width hasn't been set we use the default value.
            None => 64,
        }
    }

    // Convert the height of a cell from character units to pixels. If the
    // height hasn't been set by the user we use the default value.
    fn row_pixel_height(&mut self, row: RowNum, position: ObjectMovement) -> u32 {
        match self.changed_rows.get(&row) {
            Some(row_options) => {
                let hidden = row_options.hidden;

                if hidden && position != ObjectMovement::MoveAndSizeWithCellsAfter {
                    // A hidden row is treated as having a height of zero unless
                    // the "object_movement" is MoveAndSizeWithCellsAfter.
                    0u32
                } else {
                    (row_options.height * 4.0 / 3.0) as u32
                }
            }
            None => (self.user_default_row_height * 4.0 / 3.0) as u32,
        }
    }

    // Reset an worksheet global data or structures between saves.
    pub(crate) fn reset(&mut self) {
        xmlwriter::reset(&mut self.writer);
        xmlwriter::reset(&mut self.drawing.writer);

        for chart in self.charts.values_mut() {
            xmlwriter::reset(&mut chart.writer);
        }

        for table in &mut self.tables {
            xmlwriter::reset(&mut table.writer);
        }

        self.rel_count = 0;
        self.comment_relationships.clear();
        self.drawing_object_relationships.clear();
        self.drawing_rel_ids.clear();
        self.drawing_relationships.clear();
        self.drawing.drawings.clear();
        self.header_footer_vml_info.clear();
        self.hyperlink_relationships.clear();
        self.table_relationships.clear();
        self.vml_drawing_relationships.clear();
        self.background_relationships.clear();
    }

    // Check if any external relationships are required.
    pub(crate) fn has_relationships(&self) -> bool {
        !self.hyperlink_relationships.is_empty()
            || !self.drawing_object_relationships.is_empty()
            || !self.table_relationships.is_empty()
            || !self.background_relationships.is_empty()
    }

    // Check if there is a header image.
    pub(crate) fn has_header_footer_images(&self) -> bool {
        self.header_footer_images[0].is_some()
            || self.header_footer_images[1].is_some()
            || self.header_footer_images[2].is_some()
            || self.header_footer_images[3].is_some()
            || self.header_footer_images[4].is_some()
            || self.header_footer_images[5].is_some()
    }

    // Check that there is a header/footer &[Picture] variable in the correct
    // position to match the corresponding image object.
    fn verify_header_footer_image(string: &str, position: &HeaderImagePosition) -> bool {
        match position {
            HeaderImagePosition::Left => {
                let segments: Vec<&str> = string.split("&L").collect();
                if segments.len() == 2 {
                    let right_segment = segments[1];
                    let segments: Vec<&str> = right_segment.split("&C").collect();
                    let left_segment = segments[0];

                    let segments: Vec<&str> = left_segment.split("&R").collect();
                    let left_segment = segments[0];

                    left_segment.contains("&[Picture]") || left_segment.contains("&G")
                } else {
                    false
                }
            }
            HeaderImagePosition::Right => {
                let segments: Vec<&str> = string.split("&R").collect();
                if segments.len() == 2 {
                    let right_segment = segments[1];
                    let segments: Vec<&str> = right_segment.split("&C").collect();
                    let left_segment = segments[0];

                    let segments: Vec<&str> = left_segment.split("&L").collect();
                    let left_segment = segments[0];

                    left_segment.contains("&[Picture]") || left_segment.contains("&G")
                } else {
                    false
                }
            }
            HeaderImagePosition::Center => {
                let segments: Vec<&str> = string.split("&C").collect();
                if segments.len() == 2 {
                    let right_segment = segments[1];
                    let segments: Vec<&str> = right_segment.split("&L").collect();
                    let left_segment = segments[0];

                    let segments: Vec<&str> = left_segment.split("&R").collect();
                    let left_segment = segments[0];

                    left_segment.contains("&[Picture]") || left_segment.contains("&G")
                } else {
                    false
                }
            }
        }
    }

    // Convert column pixel width to character width.
    pub(crate) fn pixels_to_width(pixels: u16) -> f64 {
        // Properties for Calibri 11.
        let max_digit_width = 7.0_f64;
        let padding = 5.0_f64;
        let mut width = f64::from(pixels);

        if width < 12.0 {
            width /= max_digit_width + padding;
        } else {
            width = (width - padding) / max_digit_width;
        }

        width
    }

    // Return a range of data from a worksheet to use as cache data in a chart
    // file. If the range doesn't contain string/number data then we return a
    // default struct with an empty cache.
    pub(crate) fn get_cache_data(
        &self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> ChartRangeCacheData {
        let mut cache = ChartRangeCacheData::new();
        let mut data = vec![];

        for row_num in first_row..=last_row {
            match self.data_table.get(&row_num) {
                Some(columns) => {
                    for col_num in first_col..=last_col {
                        match columns.get(&col_num) {
                            Some(cell) => match cell {
                                CellType::String { string, .. } => {
                                    data.push(string.to_string());
                                    cache.cache_type = ChartRangeCacheDataType::String;
                                }
                                CellType::Number { number, .. } => {
                                    if cache.cache_type != ChartRangeCacheDataType::String {
                                        cache.cache_type = ChartRangeCacheDataType::Number;
                                    }
                                    data.push(number.to_string());
                                }
                                CellType::DateTime { number, .. } => {
                                    if cache.cache_type != ChartRangeCacheDataType::String {
                                        cache.cache_type = ChartRangeCacheDataType::Date;
                                    }
                                    data.push(number.to_string());
                                }

                                _ => data.push(String::new()),
                            },
                            None => data.push(String::new()),
                        }
                    }
                }
                None => data.push(String::new()),
            }
        }

        cache.data = data;
        cache
    }

    // Get the default header names for a worksheet table. These are generally
    // "Column1", "Column2", etc., unless the user has already specified a
    // string in the cell that will contain the header using
    // `worksheet.write()`.
    pub(crate) fn default_table_headers(
        &self,
        first_row: RowNum,
        first_col: ColNum,
        last_col: ColNum,
        show_header_row: bool,
    ) -> Vec<String> {
        let mut headers = vec![];

        for col_num in first_col..=last_col {
            headers.push(format!("Column{}", col_num - first_col + 1));
        }

        if !show_header_row {
            return headers;
        }

        if let Some(columns) = self.data_table.get(&first_row) {
            for col_num in first_col..=last_col {
                if let Some(CellType::String { string, .. }) = columns.get(&col_num) {
                    headers[(col_num - first_col) as usize] = string.to_string();
                } else if let Some(CellType::InlineString { string, .. }) = columns.get(&col_num) {
                    headers[(col_num - first_col) as usize] = string.to_string();
                }
            }
        }

        headers
    }

    // Autofit the worksheet column widths, approximately.
    //
    // Auto-fitting of columns is something that Excel does at runtime when it
    // has access to all of the worksheet information as well as the Windows
    // functions for calculating display areas based on fonts and formatting.
    //
    // The `rust_xlsxwriter` library doesn't have access to the Windows
    // functions that Excel has so it simulates autofit by calculating string
    // widths using metrics taken from Excel.
    //
    // This internal function supports autofitting to Excel's maximum cell width
    // or to a user defined value.
    fn autofit_worksheet(&mut self, max_autofit_width: u16) -> &mut Worksheet {
        let mut max_widths: HashMap<ColNum, u16> = HashMap::new();

        let (first_row, last_row) = if self.use_constant_memory {
            (self.current_row, self.current_row)
        } else {
            (self.dimensions.first_row, self.dimensions.last_row)
        };

        // Iterate over all of the data in the worksheet and find the max data
        // width for each column.
        for row_num in first_row..=last_row {
            if let Some(columns) = self.data_table.get(&row_num) {
                for col_num in self.dimensions.first_col..=self.dimensions.last_col {
                    if let Some(cell) = columns.get(&col_num) {
                        let mut pixel_width = match cell {
                            // For strings we do a calculation based on
                            // character widths taken from Excel. For rich
                            // strings we use the unformatted string. We also
                            // split multi-line strings and handle each part
                            // separately.
                            CellType::String { string, .. }
                            | CellType::InlineString { string, .. }
                            | CellType::RichString {
                                raw_string: string, ..
                            } => {
                                let mut max = 0;
                                for segment in string.lines() {
                                    let length = utility::pixel_width(segment);
                                    max = cmp::max(max, length);
                                }
                                max
                            }

                            // For numbers we use a workaround/optimization
                            // since digits all have a pixel width of 7. This
                            // gives a slightly greater width for the decimal
                            // place and minus sign but only by a few pixels and
                            // over-estimation is okay.
                            CellType::Number { number, .. } => 7 * number.to_string().len() as u16,

                            // For Boolean types we use the Excel standard
                            // widths for TRUE and FALSE.
                            CellType::Boolean { boolean, .. } => {
                                if *boolean {
                                    31
                                } else {
                                    36
                                }
                            }

                            // For formulas we autofit the result of the formula
                            // if it has a non-zero/default value.
                            CellType::Formula { result, .. }
                            | CellType::ArrayFormula { result, .. } => {
                                if result.as_ref() == "0" || result.is_empty() {
                                    0
                                } else {
                                    utility::pixel_width(result)
                                }
                            }

                            // Datetimes are just numbers but they also have an
                            // Excel format. It isn't feasible to parse the
                            // number format to get the actual string width for
                            // all format types so we use a width based on the
                            // Excel's default format: mm/dd/yyyy.
                            CellType::DateTime { .. } => 68,

                            // Ignore the following types which don't add to the width.
                            CellType::Blank { .. } | CellType::Error { .. } => 0,
                        };

                        // If the cell is in an autofilter header we add an
                        // additional 16 pixels for the dropdown arrow.
                        let is_autofilter_row = if self.use_constant_memory {
                            self.cells_with_autofilter
                                .contains_key(&(self.current_row, col_num))
                        } else {
                            self.cells_with_autofilter.contains_key(&(row_num, col_num))
                        };

                        // Add autofilter dropdown padding.
                        if pixel_width > 0 && is_autofilter_row {
                            pixel_width += 16;
                        }

                        // Add standard cell 7 pixel padding.
                        if pixel_width > 0 {
                            pixel_width += 7;
                        }

                        // Limit the autofit width to Excel's limit for `autofit()`
                        // or to a user defined value for `autofit_to_max_width()`.
                        pixel_width = std::cmp::min(pixel_width, max_autofit_width);

                        // Update the max column width.
                        if pixel_width > 0 {
                            match max_widths.get_mut(&col_num) {
                                // Update the max for the column.
                                Some(max) => {
                                    if pixel_width > *max {
                                        *max = pixel_width;
                                    };
                                }
                                None => {
                                    // Add a new column entry and maximum.
                                    max_widths.insert(col_num, pixel_width);
                                }
                            }
                        }
                    }
                }
            }
        }

        // Set the max character width for each column.
        for (col, pixels) in &max_widths {
            let width = Self::pixels_to_width(*pixels);
            self.store_column_width(*col, width, true);
        }

        self
    }

    /// Set the row properties (outline level and hidden) for a range of grouped
    /// rows in an outline.
    fn set_grouped_rows(
        &mut self,
        first_row: RowNum,
        last_row: RowNum,
        collapsed: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Get a suitable minimum value for the dimension check/set.
        let min_col = self.get_min_col();

        // Check rows are in the allowed range.
        if !self.check_dimensions(first_row, min_col) || !self.check_dimensions(last_row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_row > last_row {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Write group outline for each row.
        for row in first_row..=last_row {
            self.set_grouped_row(row, collapsed)?;
        }

        Ok(self)
    }

    // Set the row properties (outline level and hidden) for a single row in an
    // outline group. Rows are hidden if they are part of a collapsed group.
    fn set_grouped_row(
        &mut self,
        row: RowNum,
        collapsed: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Update an existing row metadata object or create a new one.
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => {
                let next_level = row_options.level + 1;

                // Check the outline level is within Excel's limits.
                if next_level > 7 {
                    return Err(XlsxError::MaxGroupLevelExceeded);
                }

                // Store max worksheet outline row level.
                self.max_outline_row_level = std::cmp::max(next_level, self.max_outline_row_level);

                row_options.level = next_level;

                if collapsed {
                    row_options.hidden = true;
                }
            }
            None => {
                let row_options = RowOptions {
                    height: self.user_default_row_height,
                    xf_index: 0,
                    hidden: collapsed,
                    level: 1,
                    collapsed: false,
                    format: None,
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    // Set the collapsed property to the row after the collapse group.
    fn set_collapsed_row(&mut self, row: RowNum) -> Result<&mut Worksheet, XlsxError> {
        // Check rows are in the allowed range.
        let min_col = self.get_min_col();
        if !self.check_dimensions(row, min_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing row metadata object or create a new one.
        match self.changed_rows.get_mut(&row) {
            Some(row_options) => row_options.collapsed = true,
            None => {
                let row_options = RowOptions {
                    height: self.user_default_row_height,
                    xf_index: 0,
                    hidden: false,
                    level: 0,
                    collapsed: true,
                    format: None,
                };
                self.changed_rows.insert(row, row_options);
            }
        }

        Ok(self)
    }

    /// Set the column properties (outline level and hidden) for a range of grouped
    /// columns in an outline.
    fn set_grouped_columns(
        &mut self,
        first_col: ColNum,
        last_col: ColNum,
        collapsed: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Get a suitable minimum value for the dimension check/set.
        let min_row = self.get_min_row();

        // Check columns are in the allowed range.
        if !self.check_dimensions(min_row, first_col) || !self.check_dimensions(min_row, last_col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Check order of first/last values.
        if first_col > last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // Write group outline for each column.
        for col in first_col..=last_col {
            self.set_grouped_column(col, collapsed)?;
        }

        Ok(self)
    }

    // Set the column properties (outline level and hidden) for a single column in an
    // outline group. Rows are hidden if they are part of a collapsed group.
    fn set_grouped_column(
        &mut self,
        column: ColNum,
        collapsed: bool,
    ) -> Result<&mut Worksheet, XlsxError> {
        // Update an existing column metadata object or create a new one.
        match self.changed_cols.get_mut(&column) {
            Some(column_options) => {
                let next_level = column_options.level + 1;

                // Check the outline level is within Excel's limits.
                if next_level > 7 {
                    return Err(XlsxError::MaxGroupLevelExceeded);
                }

                // Store max worksheet outline column level.
                self.max_outline_col_level = std::cmp::max(next_level, self.max_outline_col_level);

                column_options.level = next_level;

                if collapsed {
                    column_options.hidden = true;
                }
            }
            None => {
                let col_options = ColOptions {
                    width: DEFAULT_COL_WIDTH,
                    xf_index: 0,
                    hidden: collapsed,
                    level: 1,
                    collapsed: false,
                    autofit: false,
                    format: None,
                };
                self.changed_cols.insert(column, col_options);
            }
        }

        Ok(self)
    }

    // Set the collapsed property to the column after the collapse group.
    fn set_collapsed_column(&mut self, col: ColNum) -> Result<&mut Worksheet, XlsxError> {
        // Check columns are in the allowed range.
        let min_row = self.get_min_row();
        if !self.check_dimensions(min_row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        // Update an existing column metadata object or create a new one.
        match self.changed_cols.get_mut(&col) {
            Some(column_options) => column_options.collapsed = true,
            None => {
                let column_options = ColOptions {
                    width: DEFAULT_COL_WIDTH,
                    xf_index: 0,
                    hidden: false,
                    level: 0,
                    collapsed: true,
                    autofit: false,
                    format: None,
                };
                self.changed_cols.insert(col, column_options);
            }
        }

        Ok(self)
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file. It is split into sections that are
    // before, during and after <sheetData> to allow those sections to be
    // written separately, and to different in-memory and file based buffers, in
    // constant memory mode.
    pub(crate) fn assemble_xml_file(&mut self) {
        if self.is_chartsheet {
            return self.assemble_chartsheet();
        }

        // Write worksheet up to <sheetData>
        self.assemble_xml_file_start();

        // Write the sheetData data.
        self.write_data_table();

        // Write worksheet after <sheetData>
        self.assemble_xml_file_end();
    }

    // Assemble and write the XML file up to the <sheetData> section.
    pub(crate) fn assemble_xml_file_start(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the worksheet element.
        self.write_worksheet();

        // Write the sheetPr element.
        self.write_sheet_pr();

        // Write the dimension element.
        self.write_dimension();

        // Write the sheetViews element.
        self.write_sheet_views();

        // Write the sheetFormatPr element.
        self.write_sheet_format_pr();

        // Write the cols element.
        self.write_cols();

        // Write the <sheetData> element.
        if !self.has_sheet_data
            && self.data_table.is_empty()
            && self.notes.is_empty()
            && self.changed_rows.is_empty()
        {
            xml_empty_tag_only(&mut self.writer, "sheetData");
        } else {
            xml_start_tag_only(&mut self.writer, "sheetData");
            self.has_sheet_data = true;
        }
    }

    // Assemble and write the XML file after the <sheetData> section.
    pub(crate) fn assemble_xml_file_end(&mut self) {
        // Write the <sheetData> element.
        if self.has_sheet_data {
            xml_end_tag(&mut self.writer, "sheetData");
        }

        // Write the sheetProtection element.
        if self.protection_on {
            self.write_sheet_protection();
        }

        // Write the protectedRange element.
        if !self.unprotected_ranges.is_empty() {
            self.write_protected_ranges();
        }

        // Write the autoFilter element.
        if !self.autofilter_area.is_empty() {
            self.write_auto_filter();
        }

        // Write the mergeCells element.
        if !self.merged_ranges.is_empty() {
            self.write_merge_cells();
        }

        // Write the conditionalFormatting element.
        if self.has_conditional_formats {
            self.write_conditional_formats();
        }

        // Write the <dataValidations element.
        if !self.data_validations.is_empty() {
            self.write_data_validations();
        }

        // Write the hyperlinks elements.
        if !self.hyperlinks.is_empty() {
            self.write_hyperlinks();
        }

        // Write the printOptions element.
        if self.print_options_changed {
            self.write_print_options();
        }

        // Write the pageMargins element.
        self.write_page_margins();

        // Write the pageSetup element.
        if self.page_setup_changed {
            self.write_page_setup();
        }

        // Write the headerFooter element.
        if self.head_footer_changed {
            self.write_header_footer();
        }

        // Write the rowBreaks element.
        if !self.horizontal_breaks.is_empty() {
            self.write_row_breaks();
        }

        // Write the colBreaks element.
        if !self.vertical_breaks.is_empty() {
            self.write_col_breaks();
        }

        // Write the ignoredErrors element.
        if !self.ignored_errors.is_empty() {
            self.write_ignored_errors();
        }

        // Write the drawing element.
        if !self.drawing.drawings.is_empty() {
            self.write_drawing();
        }

        // Write the legacyDrawing element.
        if self.has_vml {
            self.write_legacy_drawing();
        }

        // Write the legacyDrawingHF element.
        if self.has_header_footer_images() {
            self.write_legacy_drawing_hf();
        }

        // Write the picture element.
        if self.background_image.is_some() {
            self.write_picture();
        }

        // Write the tableParts element.
        if !self.tables.is_empty() {
            self.write_table_parts();
        }

        // Write the extLst element.
        if self.use_x14_extensions {
            self.write_extensions();
        }

        // Close the worksheet tag.
        xml_end_tag(&mut self.writer, "worksheet");
    }

    // Assemble and write the XML file for chartsheets.
    pub(crate) fn assemble_chartsheet(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the chartsheet element.
        self.write_chartsheet();

        // Write the sheetPr element.
        self.write_sheet_pr();

        // Write the sheetViews element.
        self.write_sheet_views();

        // Write the sheetProtection element.
        if self.protection_on {
            self.write_sheet_protection();
        }

        // Write the protectedRange element.
        if !self.unprotected_ranges.is_empty() {
            self.write_protected_ranges();
        }

        // Write the printOptions element.
        if self.print_options_changed {
            self.write_print_options();
        }

        // Write the pageMargins element.
        self.write_page_margins();

        // Write the pageSetup element.
        if self.page_setup_changed {
            self.write_page_setup();
        }

        // Write the headerFooter element.
        if self.head_footer_changed {
            self.write_header_footer();
        }

        // Write the drawing element.
        if !self.drawing.drawings.is_empty() {
            self.write_drawing();
        }

        // Write the legacyDrawingHF element.
        if self.has_header_footer_images() {
            self.write_legacy_drawing_hf();
        }

        // Close the worksheet tag.
        xml_end_tag(&mut self.writer, "chartsheet");
    }

    // Write the <worksheet> element.
    fn write_worksheet(&mut self) {
        let mut attributes = vec![
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
        ];

        if self.use_x14_extensions {
            attributes.push((
                "xmlns:mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
            ));
            attributes.push((
                "xmlns:x14ac",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
            ));
            attributes.push(("mc:Ignorable", "x14ac"));
        }

        xml_start_tag(&mut self.writer, "worksheet", &attributes);
    }

    // Write the <chartsheet> element.
    fn write_chartsheet(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
        ];

        xml_start_tag(&mut self.writer, "chartsheet", &attributes);
    }

    // Write the <sheetPr> element.
    fn write_sheet_pr(&mut self) {
        if self.filter_conditions.is_empty()
            && !self.fit_to_page
            && (self.tab_color == Color::Default || self.tab_color == Color::Automatic)
            && self.vba_codename.is_none()
            && !self.is_chartsheet
            && !self.outline_symbols_above
            && !self.outline_symbols_left
        {
            return;
        }

        let mut attributes = vec![];
        if let Some(codename) = &self.vba_codename {
            attributes.push(("codeName", codename.clone()));
        }
        if !self.filter_conditions.is_empty() {
            attributes.push(("filterMode", "1".to_string()));
        }

        if self.fit_to_page
            || self.outline_symbols_above
            || self.outline_symbols_left
            || (self.tab_color != Color::Default && self.tab_color != Color::Automatic)
        {
            xml_start_tag(&mut self.writer, "sheetPr", &attributes);

            // Write the pageSetUpPr element.
            self.write_page_set_up_pr();

            // Write the tabColor element.
            self.write_tab_color();

            // Write the outlinePr element.
            self.write_outline_pr();

            xml_end_tag(&mut self.writer, "sheetPr");
        } else {
            xml_empty_tag(&mut self.writer, "sheetPr", &attributes);
        }
    }

    // Write the <pageSetUpPr> element.
    fn write_page_set_up_pr(&mut self) {
        if !self.fit_to_page {
            return;
        }

        let attributes = [("fitToPage", "1")];

        xml_empty_tag(&mut self.writer, "pageSetUpPr", &attributes);
    }

    // Write the <tabColor> element.
    fn write_tab_color(&mut self) {
        if self.tab_color == Color::Default {
            return;
        }

        let attributes = self.tab_color.attributes();

        xml_empty_tag(&mut self.writer, "tabColor", &attributes);
    }

    // Write the <outlinePr> element.
    fn write_outline_pr(&mut self) {
        if !self.outline_symbols_above && !self.outline_symbols_left {
            return;
        }

        let mut attributes = vec![];

        if self.outline_symbols_above {
            attributes.push(("summaryBelow", "0".to_string()));
        }

        if self.outline_symbols_left {
            attributes.push(("summaryRight", "0".to_string()));
        }

        xml_empty_tag(&mut self.writer, "outlinePr", &attributes);
    }

    // Write the <dimension> element.
    fn write_dimension(&mut self) {
        let mut attributes = vec![];
        let mut range = "A1".to_string();

        // Special case where all cols have been formatted but no row data.
        if self.dimensions.first_row == 0
            && self.dimensions.first_col == 0
            && self.dimensions.last_row == 0
            && self.dimensions.last_col == COL_MAX - 1
        {
            range = "A1".to_string();
        }
        // Must common case: some cells have been formatted.
        else if self.dimensions.first_row != ROW_MAX
            || self.dimensions.first_col != COL_MAX
            || self.dimensions.last_row != 0
            || self.dimensions.last_col != 0
        {
            range = utility::cell_range(
                self.dimensions.first_row,
                self.dimensions.first_col,
                self.dimensions.last_row,
                self.dimensions.last_col,
            );
        }

        attributes.push(("ref", range));

        xml_empty_tag(&mut self.writer, "dimension", &attributes);
    }

    // Write the <sheetViews> element.
    fn write_sheet_views(&mut self) {
        xml_start_tag_only(&mut self.writer, "sheetViews");

        // Write the sheetView element.
        self.write_sheet_view();

        xml_end_tag(&mut self.writer, "sheetViews");
    }

    // Write the <sheetView> element.
    fn write_sheet_view(&mut self) {
        let mut attributes = vec![];

        if !self.screen_gridlines {
            attributes.push(("showGridLines", "0".to_string()));
        }

        if self.right_to_left {
            attributes.push(("rightToLeft", "1".to_string()));
        }

        if self.selected {
            attributes.push(("tabSelected", "1".to_string()));
        }

        match self.page_view {
            PageView::PageLayout => {
                attributes.push(("view", "pageLayout".to_string()));
            }
            PageView::PageBreaks => {
                attributes.push(("view", "pageBreakPreview".to_string()));
            }
            PageView::Normal => {}
        }

        if !self.top_left_cell.is_empty() {
            attributes.push(("topLeftCell", self.top_left_cell.clone()));
        }

        if self.zoom != 100 {
            attributes.push(("zoomScale", self.zoom.to_string()));

            match self.page_view {
                PageView::PageLayout => {
                    attributes.push(("zoomScalePageLayoutView", self.zoom.to_string()));
                }
                PageView::PageBreaks => {
                    attributes.push(("zoomScaleSheetLayoutView", self.zoom.to_string()));
                }
                PageView::Normal => {
                    if !self.is_chartsheet {
                        attributes.push(("zoomScaleNormal", self.zoom.to_string()));
                    }
                }
            }
        }

        attributes.push(("workbookViewId", "0".to_string()));

        if self.panes.is_empty() && self.selected_range.0.is_empty() {
            xml_empty_tag(&mut self.writer, "sheetView", &attributes);
        } else {
            xml_start_tag(&mut self.writer, "sheetView", &attributes);
            self.write_panes();
            self.write_selections();
            xml_end_tag(&mut self.writer, "sheetView");
        }
    }

    // Write the elements associated with cell selections.
    fn write_selections(&mut self) {
        if self.selected_range.0.is_empty() {
            return;
        }

        let active_cell = self.selected_range.0.clone();
        let range = self.selected_range.1.clone();

        self.write_selection("", &active_cell, &range);
    }

    // Write the elements associated with panes.
    fn write_panes(&mut self) {
        if self.panes.is_empty() {
            return;
        }

        let row = self.panes.freeze_cell.0;
        let col = self.panes.freeze_cell.1;

        // Write the pane and selection elements.
        if row > 0 && col > 0 {
            self.write_pane("bottomRight");
            self.write_selection(
                "topRight",
                &utility::row_col_to_cell(0, col),
                &utility::row_col_to_cell(0, col),
            );
            self.write_selection(
                "bottomLeft",
                &utility::row_col_to_cell(row, 0),
                &utility::row_col_to_cell(row, 0),
            );
            self.write_selection("bottomRight", "", "");
        } else if col > 0 {
            self.write_pane("topRight");
            self.write_selection("topRight", "", "");
        } else {
            self.write_pane("bottomLeft");
            self.write_selection("bottomLeft", "", "");
        }
    }

    // Write the <pane> element.
    fn write_pane(&mut self, active_pane: &str) {
        let row = self.panes.freeze_cell.0;
        let col = self.panes.freeze_cell.1;
        let mut attributes = vec![];

        if col > 0 {
            attributes.push(("xSplit", col.to_string()));
        }

        if row > 0 {
            attributes.push(("ySplit", row.to_string()));
        }

        attributes.push(("topLeftCell", self.panes.top_left()));
        attributes.push(("activePane", active_pane.to_string()));
        attributes.push(("state", "frozen".to_string()));

        xml_empty_tag(&mut self.writer, "pane", &attributes);
    }

    // Write the <selection> element.
    fn write_selection(&mut self, position: &str, active_cell: &str, range: &str) {
        let mut attributes = vec![];

        if !position.is_empty() {
            attributes.push(("pane", position.to_string()));
        }

        if !active_cell.is_empty() {
            attributes.push(("activeCell", active_cell.to_string()));
        }

        if !range.is_empty() {
            attributes.push(("sqref", range.to_string()));
        }

        xml_empty_tag(&mut self.writer, "selection", &attributes);
    }

    // Write the <sheetFormatPr> element.
    fn write_sheet_format_pr(&mut self) {
        let mut attributes = vec![("defaultRowHeight", self.user_default_row_height.to_string())];

        if self.user_default_row_height != DEFAULT_ROW_HEIGHT {
            attributes.push(("customHeight", "1".to_string()));
        }

        if self.hide_unused_rows {
            attributes.push(("zeroHeight", "1".to_string()));
        }

        if self.max_outline_row_level > 0 {
            attributes.push(("outlineLevelRow", self.max_outline_row_level.to_string()));
        }

        if self.max_outline_col_level > 0 {
            attributes.push(("outlineLevelCol", self.max_outline_col_level.to_string()));
        }

        if self.use_x14_extensions {
            attributes.push(("x14ac:dyDescent", "0.25".to_string()));
        }

        xml_empty_tag(&mut self.writer, "sheetFormatPr", &attributes);
    }

    // Write the <mergeCells> element.
    fn write_merge_cells(&mut self) {
        let attributes = [("count", self.merged_ranges.len().to_string())];

        xml_start_tag(&mut self.writer, "mergeCells", &attributes);

        for merge_range in &self.merged_ranges.clone() {
            // Write the mergeCell element.
            self.write_merge_cell(merge_range);
        }

        xml_end_tag(&mut self.writer, "mergeCells");
    }

    // Write the <mergeCell> element.
    fn write_merge_cell(&mut self, merge_range: &CellRange) {
        let attributes = [("ref", merge_range.to_range_string())];

        xml_empty_tag(&mut self.writer, "mergeCell", &attributes);
    }

    // Write the <hyperlinks> element.
    fn write_hyperlinks(&mut self) {
        xml_start_tag_only(&mut self.writer, "hyperlinks");

        let mut rel_id = 1u32;
        for (cell, hyperlink) in &mut self.hyperlinks.clone() {
            rel_id = hyperlink.increment_rel_id(rel_id);
            self.write_hyperlink(cell.0, cell.1, hyperlink);
        }

        self.rel_count = rel_id - 1;

        xml_end_tag(&mut self.writer, "hyperlinks");
    }

    // Write the <conditionalFormatting> element.
    fn write_conditional_formats(&mut self) {
        let mut guid_index = 1;
        let mut priority = 1;

        for cell_range in &self.conditional_format_order {
            let conditionals_for_range = self.conditional_formats.get(cell_range).unwrap();

            let has_x14_only = conditionals_for_range
                .iter()
                .all(|rule| rule.has_x14_only());

            // Don't write classic CFs if range only contains new style x14 CFs.
            if !has_x14_only {
                let attributes = [("sqref", cell_range.as_str())];
                xml_start_tag(&mut self.writer, "conditionalFormatting", &attributes);
            }

            for conditional_format in conditionals_for_range {
                // Create a pseudo GUID for each unique Excel 2010 data bar.
                let mut guid = String::new();
                if conditional_format.has_x14_extensions() {
                    guid = format!(
                        "{{DA7ABA51-AAAA-BBBB-{:04X}-{:012X}}}",
                        self.sheet_index + 1,
                        guid_index
                    );
                    guid_index += 1;
                }

                // Get the format dxf_index as a global value.
                let mut dxf_index: Option<u32> = None;
                if let Some(local_index) = conditional_format.format_index() {
                    dxf_index = Some(self.global_dxf_indices[local_index as usize]);
                }

                if !conditional_format.has_x14_only() {
                    let rule = conditional_format.rule(dxf_index, priority, cell_range, &guid);
                    xml_raw_string(&mut self.writer, &rule);
                }

                priority += 1;
            }

            if !has_x14_only {
                xml_end_tag(&mut self.writer, "conditionalFormatting");
            }
        }
    }

    // Write the <x14:conditionalFormattings> element. This is used for new
    // style conditional formats added after the original spec.
    fn write_conditional_formattings(&mut self) {
        xml_start_tag_only(&mut self.writer, "x14:conditionalFormattings");

        let mut guid_index = 1;
        let mut priority = 1;

        for (cell_range, conditionals_for_range) in &self.conditional_formats {
            // Only create an entry if range contains a x14 style conditional.
            if !conditionals_for_range
                .iter()
                .any(|rule| rule.has_x14_extensions())
            {
                continue;
            }

            let attributes = [(
                "xmlns:xm",
                "http://schemas.microsoft.com/office/excel/2006/main",
            )];

            xml_start_tag(&mut self.writer, "x14:conditionalFormatting", &attributes);

            for conditional_format in conditionals_for_range {
                if conditional_format.has_x14_extensions() {
                    // Create a pseudo GUID for each unique Excel 2010 data bar.
                    let guid = format!(
                        "{{DA7ABA51-AAAA-BBBB-{:04X}-{:012X}}}",
                        self.sheet_index + 1,
                        guid_index
                    );
                    guid_index += 1;

                    let rule = conditional_format.x14_rule(priority, &guid);
                    xml_raw_string(&mut self.writer, &rule);
                }
                priority += 1;
            }

            xml_data_element_only(&mut self.writer, "xm:sqref", cell_range);
            xml_end_tag(&mut self.writer, "x14:conditionalFormatting");
        }

        xml_end_tag(&mut self.writer, "x14:conditionalFormattings");
    }

    // Write the <dataValidations> element.
    fn write_data_validations(&mut self) {
        let attributes = [("count", self.data_validations.len().to_string())];

        xml_start_tag(&mut self.writer, "dataValidations", &attributes);

        for (range, data_validation) in &self.data_validations.clone() {
            // Write the dataValidation element.
            self.write_data_validation(range, data_validation);
        }

        xml_end_tag(&mut self.writer, "dataValidations");
    }

    // Write the <dataValidation> element.
    fn write_data_validation(&mut self, range: &String, data_validation: &DataValidation) {
        // The Any type doesn't have a rule or values so handle that separately.
        if data_validation.validation_type == DataValidationType::Any {
            self.write_data_validation_any(range, data_validation);
            return;
        }

        // Start the attributes.
        let mut attributes = vec![("type", data_validation.validation_type.to_string())];

        match data_validation.error_style {
            DataValidationErrorStyle::Warning | DataValidationErrorStyle::Information => {
                attributes.push(("errorStyle", data_validation.error_style.to_string()));
            }
            DataValidationErrorStyle::Stop => {}
        }

        match &data_validation.rule {
            &DataValidationRuleInternal::Between(_, _)
            | DataValidationRuleInternal::CustomFormula(_)
            | DataValidationRuleInternal::ListSource(_) => {
                // Excel doesn't use an operator for these types.
            }
            _ => {
                attributes.push(("operator", data_validation.rule.to_string()));
            }
        };

        if data_validation.ignore_blank {
            attributes.push(("allowBlank", "1".to_string()));
        }

        if !data_validation.show_dropdown {
            attributes.push(("showDropDown", "1".to_string()));
        }

        if data_validation.show_input_message {
            attributes.push(("showInputMessage", "1".to_string()));
        }

        if data_validation.show_error_message {
            attributes.push(("showErrorMessage", "1".to_string()));
        }

        if !data_validation.error_title.is_empty() {
            attributes.push(("errorTitle", data_validation.error_title.clone()));
        }

        if !data_validation.error_message.is_empty() {
            attributes.push(("error", data_validation.error_message.clone()));
        }

        if !data_validation.input_title.is_empty() {
            attributes.push(("promptTitle", data_validation.input_title.clone()));
        }

        if !data_validation.input_message.is_empty() {
            attributes.push(("prompt", data_validation.input_message.clone()));
        }

        attributes.push(("sqref", range.to_string()));

        xml_start_tag(&mut self.writer, "dataValidation", &attributes);

        // Write the <formula1>/<formula2> elements.
        match &data_validation.rule {
            DataValidationRuleInternal::EqualTo(value)
            | DataValidationRuleInternal::NotEqualTo(value)
            | DataValidationRuleInternal::LessThan(value)
            | DataValidationRuleInternal::LessThanOrEqualTo(value)
            | DataValidationRuleInternal::GreaterThan(value)
            | DataValidationRuleInternal::GreaterThanOrEqualTo(value)
            | DataValidationRuleInternal::ListSource(value)
            | DataValidationRuleInternal::CustomFormula(value) => {
                xml_data_element_only(&mut self.writer, "formula1", value);
            }
            DataValidationRuleInternal::Between(min, max)
            | DataValidationRuleInternal::NotBetween(min, max) => {
                xml_data_element_only(&mut self.writer, "formula1", min);
                xml_data_element_only(&mut self.writer, "formula2", max);
            }
        }
        xml_end_tag(&mut self.writer, "dataValidation");
    }

    // Write the <dataValidation> element.
    fn write_data_validation_any(&mut self, range: &String, data_validation: &DataValidation) {
        let mut attributes = vec![];

        if data_validation.ignore_blank {
            attributes.push(("allowBlank", "1".to_string()));
        }

        if !data_validation.show_dropdown {
            attributes.push(("showDropDown", "1".to_string()));
        }

        if data_validation.show_input_message {
            attributes.push(("showInputMessage", "1".to_string()));
        }

        if data_validation.show_error_message {
            attributes.push(("showErrorMessage", "1".to_string()));
        }

        if !data_validation.error_title.is_empty() {
            attributes.push(("errorTitle", data_validation.error_title.clone()));
        }

        if !data_validation.error_message.is_empty() {
            attributes.push(("error", data_validation.error_message.clone()));
        }

        if !data_validation.input_title.is_empty() {
            attributes.push(("promptTitle", data_validation.input_title.clone()));
        }

        if !data_validation.input_message.is_empty() {
            attributes.push(("prompt", data_validation.input_message.clone()));
        }

        attributes.push(("sqref", range.to_string()));

        xml_empty_tag(&mut self.writer, "dataValidation", &attributes);
    }

    // Write the <hyperlink> element.
    fn write_hyperlink(&mut self, row: RowNum, col: ColNum, hyperlink: &Url) {
        let mut attributes = vec![("ref", utility::row_col_to_cell(row, col))];

        match hyperlink.link_type {
            HyperlinkType::Url | HyperlinkType::File => {
                let rel_id = hyperlink.rel_id;
                attributes.push(("r:id", format!("rId{rel_id}")));

                if !hyperlink.rel_anchor.is_empty() {
                    attributes.push(("location", hyperlink.rel_anchor.to_string()));
                }

                if hyperlink.rel_display {
                    attributes.push(("display", hyperlink.url_link.to_string()));
                }

                if !hyperlink.tool_tip.is_empty() {
                    attributes.push(("tooltip", hyperlink.tool_tip.to_string()));
                }

                // Store the linkage to the worksheets rels file.
                self.hyperlink_relationships.push((
                    "hyperlink".to_string(),
                    hyperlink.url_link.to_string(),
                    "External".to_string(),
                ));
            }
            HyperlinkType::Internal => {
                // Internal links don't use the rel file reference id.
                attributes.push(("location", hyperlink.rel_anchor.to_string()));

                if !hyperlink.tool_tip.is_empty() {
                    attributes.push(("tooltip", hyperlink.tool_tip.to_string()));
                }

                attributes.push(("display", hyperlink.user_text.to_string()));
            }
            HyperlinkType::Unknown => {}
        }

        xml_empty_tag(&mut self.writer, "hyperlink", &attributes);
    }

    // Write the <printOptions> element.
    fn write_print_options(&mut self) {
        let mut attributes = vec![];

        if self.center_horizontally {
            attributes.push(("horizontalCentered", "1".to_string()));
        }

        if self.center_vertically {
            attributes.push(("verticalCentered", "1".to_string()));
        }

        if self.print_headings {
            attributes.push(("headings", "1".to_string()));
        }

        if self.print_gridlines {
            attributes.push(("gridLines", "1".to_string()));
        }

        xml_empty_tag(&mut self.writer, "printOptions", &attributes);
    }

    // Write the <pageMargins> element.
    fn write_page_margins(&mut self) {
        let attributes = [
            ("left", self.margin_left.to_string()),
            ("right", self.margin_right.to_string()),
            ("top", self.margin_top.to_string()),
            ("bottom", self.margin_bottom.to_string()),
            ("header", self.margin_header.to_string()),
            ("footer", self.margin_footer.to_string()),
        ];

        xml_empty_tag(&mut self.writer, "pageMargins", &attributes);
    }

    // Write the <pageSetup> element.
    fn write_page_setup(&mut self) {
        let mut attributes = vec![];

        if self.paper_size > 0 {
            attributes.push(("paperSize", self.paper_size.to_string()));
        }

        if self.print_scale != 100 {
            attributes.push(("scale", self.print_scale.to_string()));
        }

        if self.fit_to_page {
            if self.fit_width != 1 {
                attributes.push(("fitToWidth", self.fit_width.to_string()));
            }
            if self.fit_height != 1 {
                attributes.push(("fitToHeight", self.fit_height.to_string()));
            }
        }

        if !self.default_page_order {
            attributes.push(("pageOrder", "overThenDown".to_string()));
        }

        if self.is_portrait {
            attributes.push(("orientation", "portrait".to_string()));
        } else {
            attributes.push(("orientation", "landscape".to_string()));
        }

        if self.first_page_number > 0 {
            attributes.push(("useFirstPageNumber", self.first_page_number.to_string()));
        }

        if self.print_black_and_white {
            attributes.push(("blackAndWhite", "1".to_string()));
        }

        if self.print_draft {
            attributes.push(("draft", "1".to_string()));
        }

        attributes.push(("horizontalDpi", "200".to_string()));
        attributes.push(("verticalDpi", "200".to_string()));

        xml_empty_tag(&mut self.writer, "pageSetup", &attributes);
    }

    // Write the <autoFilter> element.
    fn write_auto_filter(&mut self) {
        let attributes = [("ref", self.autofilter_area.clone())];

        if self.filter_conditions.is_empty() {
            xml_empty_tag(&mut self.writer, "autoFilter", &attributes);
        } else {
            xml_start_tag(&mut self.writer, "autoFilter", &attributes);
            let col_offset = self.autofilter_defined_name.first_col;

            for col in self.filter_conditions.clone().keys() {
                let filter_condition = self.filter_conditions.get(col).unwrap().clone();

                self.write_filter_column(*col - col_offset, &filter_condition);
            }

            xml_end_tag(&mut self.writer, "autoFilter");
        }
    }

    // Write the <filterColumn> element.
    fn write_filter_column(&mut self, col: ColNum, filter_condition: &FilterCondition) {
        let attributes = [("colId", col.to_string())];

        xml_start_tag(&mut self.writer, "filterColumn", &attributes);

        if filter_condition.is_list_filter {
            self.write_list_filters(filter_condition);
        } else {
            self.write_custom_filters(filter_condition);
        }

        xml_end_tag(&mut self.writer, "filterColumn");
    }

    // Write the <filters> element.
    fn write_list_filters(&mut self, filter_condition: &FilterCondition) {
        let mut attributes = vec![];

        if filter_condition.should_match_blanks {
            attributes.push(("blank", "1".to_string()));
        }

        if filter_condition.list.is_empty() {
            xml_empty_tag(&mut self.writer, "filters", &attributes);
        } else {
            xml_start_tag(&mut self.writer, "filters", &attributes);

            for data in &filter_condition.list {
                // Write the filter element.
                self.write_filter(data.string.clone());
            }

            xml_end_tag(&mut self.writer, "filters");
        }
    }

    // Write the <filter> element.
    fn write_filter(&mut self, value: String) {
        let attributes = [("val", value)];

        xml_empty_tag(&mut self.writer, "filter", &attributes);
    }

    // Write the <customFilters> element.
    fn write_custom_filters(&mut self, filter_condition: &FilterCondition) {
        let mut attributes = vec![];

        if !filter_condition.apply_logical_or {
            attributes.push(("and", "1".to_string()));
        }

        xml_start_tag(&mut self.writer, "customFilters", &attributes);

        if let Some(data) = filter_condition.custom1.as_ref() {
            self.write_custom_filter(data);
        }
        if let Some(data) = filter_condition.custom2.as_ref() {
            self.write_custom_filter(data);
        }

        xml_end_tag(&mut self.writer, "customFilters");
    }

    // Write the <customFilter> element.
    fn write_custom_filter(&mut self, data: &FilterData) {
        let mut attributes = vec![];

        if !data.criteria.operator().is_empty() {
            attributes.push(("operator", data.criteria.operator()));
        }

        attributes.push(("val", data.value()));

        xml_empty_tag(&mut self.writer, "customFilter", &attributes);
    }

    // Store unique strings in the SST table and convert them to a string id
    // which is used when writing out the string cells.
    pub(crate) fn update_string_table_ids(&mut self, string_table: Arc<Mutex<SharedStringsTable>>) {
        if !self.has_local_string_table {
            return;
        }

        let mut string_table = string_table.lock().unwrap();

        for columns in self.data_table.values_mut() {
            for cell in columns.values_mut() {
                match cell {
                    CellType::String {
                        string, string_id, ..
                    }
                    | CellType::RichString {
                        string, string_id, ..
                    } => {
                        if string_id.is_none() {
                            let string_index = string_table.shared_string_index(Arc::clone(string));
                            *string_id = Some(string_index);
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    // Write out all the row and cell data in the worksheet data table.
    #[allow(clippy::too_many_lines)]
    fn write_data_table(&mut self) {
        let spans = self.calculate_spans();
        let mut col_names = HashMap::new();

        // Swap out the worksheet data structures so we can iterate over them and
        // still call self.write_*() methods.
        let mut temp_table: BTreeMap<RowNum, BTreeMap<ColNum, CellType>> = BTreeMap::new();
        let mut temp_changed_rows: HashMap<RowNum, RowOptions> = HashMap::new();
        mem::swap(&mut temp_table, &mut self.data_table);
        mem::swap(&mut temp_changed_rows, &mut self.changed_rows);

        for row_num in self.dimensions.first_row..=self.dimensions.last_row {
            let span_index = row_num / 16;
            let span = spans.get(&span_index).map(AsRef::as_ref);

            let row_options = temp_changed_rows.get(&row_num);
            let row_has_notes = self.notes.contains_key(&row_num);

            // If there is no column data then only the <row> metadata needs updating.
            let Some(columns) = temp_table.get(&row_num) else {
                if row_options.is_some() || row_has_notes {
                    self.write_table_row(row_num, span, row_options, false);
                }
                continue;
            };

            // The row has data. Write it out cell by cell.
            self.write_table_row(row_num, span, row_options, true);
            for (&col_num, cell) in columns {
                // Faster column name lookup for inner loop.
                let col_name = if col_num < 26 {
                    &COLUMN_LETTERS[col_num as usize..(col_num + 1) as usize]
                } else {
                    col_names
                        .entry(col_num)
                        .or_insert_with(|| utility::column_number_to_name(col_num))
                };

                match cell {
                    CellType::Number { number, xf_index }
                    | CellType::DateTime { number, xf_index } => {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        Self::write_number_cell(
                            &mut self.writer,
                            row_num + 1,
                            col_name,
                            *number,
                            xf_index,
                        );
                    }

                    CellType::String {
                        string_id,
                        xf_index,
                        ..
                    }
                    | CellType::RichString {
                        string_id,
                        xf_index,
                        ..
                    } => {
                        if let Some(string_id) = string_id {
                            let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                            Self::write_string_cell(
                                &mut self.writer,
                                row_num + 1,
                                col_name,
                                *string_id,
                                xf_index,
                            );
                        }
                    }

                    CellType::InlineString {
                        string, xf_index, ..
                    } => {
                        Self::write_inline_string_cell(
                            &mut self.writer,
                            row_num + 1,
                            col_name,
                            string,
                            *xf_index,
                        );
                    }

                    CellType::Formula {
                        formula,
                        xf_index,
                        result,
                    } => {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        Self::write_formula_cell(
                            &mut self.writer,
                            row_num + 1,
                            col_name,
                            formula,
                            xf_index,
                            result,
                        );
                    }

                    CellType::ArrayFormula {
                        formula,
                        xf_index,
                        result,
                        is_dynamic,
                        range,
                    } => {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        Self::write_array_formula_cell(
                            &mut self.writer,
                            row_num + 1,
                            col_name,
                            formula,
                            xf_index,
                            result,
                            *is_dynamic,
                            range,
                        );
                    }

                    CellType::Blank { xf_index } => {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        Self::write_blank_cell(&mut self.writer, row_num + 1, col_name, xf_index);
                    }

                    CellType::Boolean { boolean, xf_index } => {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        Self::write_boolean_cell(
                            &mut self.writer,
                            row_num + 1,
                            col_name,
                            *boolean,
                            xf_index,
                        );
                    }

                    CellType::Error { value, xf_index } => {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        let image_id = self.global_embedded_image_indices[*value as usize];
                        Self::write_error_cell(
                            &mut self.writer,
                            row_num + 1,
                            col_name,
                            image_id,
                            xf_index,
                        );
                    }
                }
            }
            xml_end_tag(&mut self.writer, "row");
        }

        // Swap back in data.
        mem::swap(&mut temp_table, &mut self.data_table);
        mem::swap(&mut temp_changed_rows, &mut self.changed_rows);
    }

    // Flush the last row of constant memory data, the write-ahead cache and any
    // modified rows.
    #[cfg(feature = "constant_memory")]
    pub(crate) fn flush_last_row(&mut self) {
        // First find any write ahead cached rows.
        let mut remaining_rows: Vec<_> = self.write_ahead.keys().copied().collect();

        // Add any changed row elements.
        let max_changed_row = self.changed_rows.keys().max().unwrap_or(&0);
        if *max_changed_row > self.current_row {
            for row_num in self.current_row..=*max_changed_row {
                if self.changed_rows.contains_key(&row_num) {
                    remaining_rows.push(row_num);
                }
            }
        }

        // Add the current row.
        remaining_rows.push(self.current_row);

        // Remove any duplicate rows.
        remaining_rows.sort_unstable();
        remaining_rows.dedup();

        // If there is data to write (and not just write-ahead or rows) we move
        // the final row ahead by one. This will start any subsequent write() on
        // the next row, after a resave.
        if self.data_table.contains_key(&self.current_row) {
            remaining_rows.push(1 + *remaining_rows.last().unwrap());
        }

        // Flush all the remaining rows.
        for remaining_row in remaining_rows {
            self.flush_data_row(remaining_row);
        }
    }

    // Flush all constant memory data up to the next target row.
    #[cfg(feature = "constant_memory")]
    fn flush_to_row(&mut self, next_row: RowNum) {
        // First find any write ahead cached rows.
        let mut intermediate_rows: Vec<_> = self
            .write_ahead
            .keys()
            .filter(|x| **x < next_row)
            .copied()
            .collect();

        // Add any changed row elements.
        for row_num in self.current_row + 1..next_row {
            if self.changed_rows.contains_key(&row_num) {
                intermediate_rows.push(row_num);
            }
        }

        // Remove any duplicate rows.
        intermediate_rows.sort_unstable();
        intermediate_rows.dedup();

        for intermediate_row in intermediate_rows {
            self.flush_data_row(intermediate_row);
        }

        self.flush_data_row(next_row);
    }

    // Write out all the row and cell data in the constant memory data table.
    #[allow(clippy::too_many_lines)]
    #[cfg(feature = "constant_memory")]
    fn flush_data_row(&mut self, next_row: RowNum) {
        let current_row = self.current_row;

        // Swap out the worksheet data structures so we can iterate over them and
        // still call self.write_*() methods.
        let mut temp_table: BTreeMap<RowNum, BTreeMap<ColNum, CellType>> = BTreeMap::new();
        let mut temp_changed_rows: HashMap<RowNum, RowOptions> = HashMap::new();
        mem::swap(&mut temp_table, &mut self.data_table);
        mem::swap(&mut temp_changed_rows, &mut self.changed_rows);

        let row_options = temp_changed_rows.get(&current_row);

        // If there is no column data then only the <row> metadata needs updating.
        let Some(columns) = temp_table.get(&current_row) else {
            if row_options.is_some() {
                self.write_constant_table_row(current_row, row_options, false);
            }

            mem::swap(&mut temp_changed_rows, &mut self.changed_rows);

            // Replace the "current" row with one from the write ahead buffer.
            // Or else it defaults to the new clean buffer created above.
            if let Some(columns) = self.write_ahead.remove(&next_row) {
                self.data_table.insert(next_row, columns);
            }
            self.current_row = next_row;

            return;
        };

        // The row has data. Write it out cell by cell.
        self.write_constant_table_row(current_row, row_options, true);
        for (&col_num, cell) in columns {
            let col_name = if col_num < 26 {
                &COLUMN_LETTERS[col_num as usize..(col_num + 1) as usize]
            } else {
                &utility::column_number_to_name(col_num)
            };

            match cell {
                CellType::Number { number, xf_index } | CellType::DateTime { number, xf_index } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    Self::write_number_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        *number,
                        xf_index,
                    );
                }

                CellType::String {
                    string_id,
                    xf_index,
                    ..
                } => {
                    if let Some(string_id) = string_id {
                        let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                        Self::write_string_cell(
                            &mut self.file_writer,
                            current_row + 1,
                            col_name,
                            *string_id,
                            xf_index,
                        );
                    }
                }

                CellType::RichString {
                    string, xf_index, ..
                } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    Self::write_inline_rich_string_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        string,
                        xf_index,
                    );
                }

                CellType::InlineString {
                    string, xf_index, ..
                } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    Self::write_inline_string_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        string,
                        xf_index,
                    );
                }

                CellType::Formula {
                    formula,
                    xf_index,
                    result,
                } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    Self::write_formula_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        formula,
                        xf_index,
                        result,
                    );
                }

                CellType::ArrayFormula {
                    formula,
                    xf_index,
                    result,
                    is_dynamic,
                    range,
                } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    Self::write_array_formula_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        formula,
                        xf_index,
                        result,
                        *is_dynamic,
                        range,
                    );
                }

                CellType::Blank { xf_index } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    Self::write_blank_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        xf_index,
                    );
                }

                CellType::Boolean { boolean, xf_index } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);

                    Self::write_boolean_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        *boolean,
                        xf_index,
                    );
                }

                CellType::Error { value, xf_index } => {
                    let xf_index = self.get_cell_xf_index(*xf_index, row_options, col_num);
                    let image_id = self.global_embedded_image_indices[*value as usize];
                    Self::write_error_cell(
                        &mut self.file_writer,
                        current_row + 1,
                        col_name,
                        image_id,
                        xf_index,
                    );
                }
            }
        }
        xml_end_tag(&mut self.file_writer, "row");

        // Swap back in changed rows data.
        mem::swap(&mut temp_changed_rows, &mut self.changed_rows);

        // Replace the "current" row with one from the write ahead buffer. Or
        // else it defaults to the new clean buffer created above.
        if let Some(columns) = self.write_ahead.remove(&next_row) {
            self.data_table.insert(next_row, columns);
        }
        self.current_row = next_row;
    }

    // Calculate the "spans" attribute of the <row> tag. This is an xlsx
    // optimization and isn't strictly required. However, it makes comparing
    // files easier. The span is the same for each block of 16 rows.
    fn calculate_spans(&mut self) -> HashMap<u32, String> {
        let mut spans: HashMap<RowNum, String> = HashMap::new();
        let mut span_min = COL_MAX;
        let mut span_max = 0;

        for row_num in self.dimensions.first_row..=self.dimensions.last_row {
            if let Some(columns) = self.data_table.get(&row_num) {
                for &col_num in columns.keys() {
                    if span_min == COL_MAX {
                        span_min = col_num;
                        span_max = col_num;
                    } else {
                        span_min = cmp::min(span_min, col_num);
                        span_max = cmp::max(span_max, col_num);
                    }
                }
            }

            if let Some(columns) = self.notes.get(&row_num) {
                for &col_num in columns.keys() {
                    if span_min == COL_MAX {
                        span_min = col_num;
                        span_max = col_num;
                    } else {
                        span_min = cmp::min(span_min, col_num);
                        span_max = cmp::max(span_max, col_num);
                    }
                }
            }

            // Store the span range for each block or 16 rows.
            if (row_num + 1) % 16 == 0 || row_num == self.dimensions.last_row {
                let span_index = row_num / 16;
                if span_min != COL_MAX {
                    span_min += 1;
                    span_max += 1;
                    let span_range = format!("{span_min}:{span_max}");
                    spans.insert(span_index, span_range);
                    span_min = COL_MAX;
                }
            }
        }

        spans
    }

    // Write the <row> element.
    fn write_table_row(
        &mut self,
        row_num: RowNum,
        span: Option<&str>,
        row_options: Option<&RowOptions>,
        has_data: bool,
    ) {
        let row_num = (row_num + 1).to_string();
        let mut attributes = vec![("r", row_num)];

        if let Some(span_range) = span {
            attributes.push(("spans", span_range.to_string()));
        }

        if self.use_x14_extensions {
            attributes.push(("x14ac:dyDescent", "0.25".to_string()));
        }

        if let Some(row_options) = row_options {
            let xf_index = row_options.xf_index;

            if xf_index != 0 {
                let xf_index = self.get_global_xf_index(xf_index);
                attributes.push(("s", xf_index.to_string()));
                attributes.push(("customFormat", "1".to_string()));
            }

            // Only add ht parameter if the height is non-default.
            if row_options.height != DEFAULT_ROW_HEIGHT
                || (row_options.height == DEFAULT_ROW_HEIGHT
                    && self.user_default_row_height != DEFAULT_ROW_HEIGHT)
            {
                attributes.push(("ht", row_options.height.to_string()));
            }

            if row_options.hidden {
                attributes.push(("hidden", "1".to_string()));
            }

            // Only add customHeight parameter if the height is non-default.
            if row_options.height != DEFAULT_ROW_HEIGHT
                || (row_options.height == DEFAULT_ROW_HEIGHT
                    && self.user_default_row_height != DEFAULT_ROW_HEIGHT)
            {
                attributes.push(("customHeight", "1".to_string()));
            }

            if row_options.level > 0 {
                attributes.push(("outlineLevel", row_options.level.to_string()));
            }

            if row_options.collapsed {
                attributes.push(("collapsed", "1".to_string()));
            }
        } else if self.user_default_row_height != DEFAULT_ROW_HEIGHT {
            attributes.push(("ht", self.user_default_row_height.to_string()));
            attributes.push(("customHeight", "1".to_string()));
        }

        if has_data {
            xml_start_tag(&mut self.writer, "row", &attributes);
        } else {
            xml_empty_tag(&mut self.writer, "row", &attributes);
        }
    }

    // Write the <row> element in constant memory mode.
    #[cfg(feature = "constant_memory")]
    fn write_constant_table_row(
        &mut self,
        row_num: RowNum,
        row_options: Option<&RowOptions>,
        has_data: bool,
    ) {
        let row_num = (row_num + 1).to_string();
        let mut attributes = vec![("r", row_num)];

        if self.use_x14_extensions {
            attributes.push(("x14ac:dyDescent", "0.25".to_string()));
        }

        if let Some(row_options) = row_options {
            let xf_index = row_options.xf_index;

            if xf_index != 0 {
                attributes.push(("s", xf_index.to_string()));
                attributes.push(("customFormat", "1".to_string()));
            }

            // Only add ht parameter if the height is non-default.
            if row_options.height != DEFAULT_ROW_HEIGHT
                || (row_options.height == DEFAULT_ROW_HEIGHT
                    && self.user_default_row_height != DEFAULT_ROW_HEIGHT)
            {
                attributes.push(("ht", row_options.height.to_string()));
            }

            if row_options.hidden {
                attributes.push(("hidden", "1".to_string()));
            }

            // Only add customHeight parameter if the height is non-default.
            if row_options.height != DEFAULT_ROW_HEIGHT
                || (row_options.height == DEFAULT_ROW_HEIGHT
                    && self.user_default_row_height != DEFAULT_ROW_HEIGHT)
            {
                attributes.push(("customHeight", "1".to_string()));
            }
        } else if self.user_default_row_height != DEFAULT_ROW_HEIGHT {
            attributes.push(("ht", self.user_default_row_height.to_string()));
            attributes.push(("customHeight", "1".to_string()));
        }

        if has_data {
            xml_start_tag(&mut self.file_writer, "row", &attributes);
        } else {
            xml_empty_tag(&mut self.file_writer, "row", &attributes);
        }
    }

    // Note: the following write_*_cell() functions are optimized for minimum
    // overhead since they are in the inner loop for all worksheet data.

    // Write the <c> element for a number.
    fn write_number_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        number: f64,
        xf_index: u32,
    ) {
        // Use the optional `ryu` crate to format f64 cell number data as a
        // string. Note, the the slightly faster `format_finite()` buffer
        // function is safe to use here since nan/inf numbers are filtered out
        // at the `store_number()` level and written as strings.
        #[cfg(feature = "ryu")]
        let mut buffer = ryu::Buffer::new();
        #[cfg(feature = "ryu")]
        let number = buffer.format_finite(number);

        if xf_index > 0 {
            write!(
                writer,
                r#"<c r="{col_name}{row}" s="{xf_index}"><v>{number}</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(writer, r#"<c r="{col_name}{row}"><v>{number}</v></c>"#,)
                .expect(XML_WRITE_ERROR);
        }
    }

    // Write the <c> element for a string.
    fn write_string_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        string_index: u32,
        xf_index: u32,
    ) {
        if xf_index > 0 {
            write!(
                writer,
                r#"<c r="{col_name}{row}" s="{xf_index}" t="s"><v>{string_index}</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(
                writer,
                r#"<c r="{col_name}{row}" t="s"><v>{string_index}</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        }
    }

    // Write the <c> element for an inline string.
    fn write_inline_string_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        string: &str,
        xf_index: u32,
    ) {
        let whitespace = ['\t', '\n', ' '];
        let preserve = if string.starts_with(whitespace) || string.ends_with(whitespace) {
            " xml:space=\"preserve\""
        } else {
            ""
        };

        let string = crate::xmlwriter::escape_xml_escapes(string);
        let string = crate::xmlwriter::escape_xml_data(&string);

        if xf_index > 0 {
            write!(
                writer,
                r#"<c r="{col_name}{row}" s="{xf_index}" t="inlineStr"><is><t{preserve}>{string}</t></is></c>"#
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(
                writer,
                r#"<c r="{col_name}{row}" t="inlineStr"><is><t{preserve}>{string}</t></is></c>"#
            )
            .expect(XML_WRITE_ERROR);
        }
    }

    // Write the <c> element for an inline rich string.
    #[cfg(feature = "constant_memory")]
    fn write_inline_rich_string_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        string: &str,
        xf_index: u32,
    ) {
        if xf_index > 0 {
            write!(
                writer,
                r#"<c r="{col_name}{row}" s="{xf_index}" t="inlineStr"><is>{string}</is></c>"#
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(
                writer,
                r#"<c r="{col_name}{row}" t="inlineStr"><is>{string}</is></c>"#
            )
            .expect(XML_WRITE_ERROR);
        }
    }

    // Write the <c> element for a formula.
    fn write_formula_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        formula: &str,
        xf_index: u32,
        result: &str,
    ) {
        let style = if xf_index > 0 {
            format!(r#" s="{xf_index}""#)
        } else {
            String::new()
        };

        // Get the result type attribute.
        let result_type = if result.parse::<f64>().is_err() {
            match result {
                // Handle error results.
                "#DIV/0!" | "#N/A" | "#NAME?" | "#NULL!" | "#NUM!" | "#REF!" | "#VALUE!"
                | "#GETTING_DATA" => r#" t="e""#,

                // Handle boolean results.
                "TRUE" | "FALSE" => r#" t="b""#,

                // Handle string results.
                _ => r#" t="str""#,
            }
        } else {
            // Handle/ignore for numeric results.
            ""
        };

        let formula = crate::xmlwriter::escape_xml_data(formula);
        let mut result = crate::xmlwriter::escape_xml_data(result);

        if result == "TRUE" {
            result = Cow::Owned("1".to_string());
        }
        if result == "FALSE" {
            result = Cow::Owned("0".to_string());
        }

        write!(
            writer,
            r#"<c r="{col_name}{row}"{style}{result_type}><f>{formula}</f><v>{result}</v></c>"#
        )
        .expect(XML_WRITE_ERROR);
    }

    // Write the <c> element for an array formula.
    #[allow(clippy::too_many_arguments)]
    fn write_array_formula_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        formula: &str,
        xf_index: u32,
        result: &str,
        is_dynamic: bool,
        range: &str,
    ) {
        let style = if xf_index > 0 {
            format!(r#" s="{xf_index}""#)
        } else {
            String::new()
        };

        let cm = if is_dynamic { r#" cm="1""# } else { "" };

        // Get the result type attribute.
        let result_type = if result.parse::<f64>().is_err() {
            match result {
                // Handle error results.
                "#DIV/0!" | "#N/A" | "#NAME?" | "#NULL!" | "#NUM!" | "#REF!" | "#VALUE!"
                | "#GETTING_DATA" => r#" t="e""#,

                // Handle boolean results.
                "TRUE" | "FALSE" => r#" t="b""#,

                // Handle string results.
                _ => r#" t="str""#,
            }
        } else {
            // Handle/ignore for numeric results.
            ""
        };

        let formula = crate::xmlwriter::escape_xml_data(formula);
        let mut result = crate::xmlwriter::escape_xml_data(result);

        if result == "TRUE" {
            result = Cow::Owned("1".to_string());
        }
        if result == "FALSE" {
            result = Cow::Owned("0".to_string());
        }

        write!(
            writer,
            r#"<c r="{col_name}{row}"{style}{cm}{result_type}><f t="array" ref="{range}">{formula}</f><v>{result}</v></c>"#
        )
        .expect(XML_WRITE_ERROR);
    }

    // Write the <c> element for a blank cell.
    fn write_blank_cell<W: Write>(writer: &mut W, row: RowNum, col_name: &str, xf_index: u32) {
        // Write formatted blank cells and ignore unformatted blank cells (like
        // Excel does).
        if xf_index > 0 {
            write!(writer, r#"<c r="{col_name}{row}" s="{xf_index}"/>"#).expect(XML_WRITE_ERROR);
        }
    }

    // Write the <c> element for a boolean cell.
    fn write_boolean_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        boolean: bool,
        xf_index: u32,
    ) {
        let boolean = i32::from(boolean);

        if xf_index > 0 {
            write!(
                writer,
                r#"<c r="{col_name}{row}" s="{xf_index}" t="b"><v>{boolean}</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(
                writer,
                r#"<c r="{col_name}{row}" t="b"><v>{boolean}</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        }
    }

    // Write the <c> element for an error cell. We currently only support the
    // #VALUE! error type which is used for embedded images.
    fn write_error_cell<W: Write>(
        writer: &mut W,
        row: RowNum,
        col_name: &str,
        value: u32,
        xf_index: u32,
    ) {
        if xf_index > 0 {
            write!(
                writer,
                r#"<c r="{col_name}{row}" s="{xf_index}" t="e" vm="{value}"><v>#VALUE!</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(
                writer,
                r#"<c r="{col_name}{row}" t="e" vm="{value}"><v>#VALUE!</v></c>"#
            )
            .expect(XML_WRITE_ERROR);
        }
    }

    // Write the <cols> element.
    fn write_cols(&mut self) {
        if self.changed_cols.is_empty() {
            return;
        }

        xml_start_tag_only(&mut self.writer, "cols");

        // We need to write contiguous equivalent columns as a range with first
        // and last columns, so we convert the HashMap to a sorted vector and
        // iterate over that.
        let mut col_options: Vec<_> = self
            .changed_cols
            .iter()
            .map(|(k, v)| (*k, v.clone()))
            .collect();
        col_options.sort_unstable_by_key(|x| x.0);
        let mut col_options = col_options.into_iter();

        // Remove the first (key, value) tuple in the vector and use it to set
        // the initial/previous properties.
        let first_col_options = col_options.next().unwrap();
        let mut first_col = first_col_options.0;
        let mut prev_col_options = first_col_options.1;
        let mut last_col = first_col;

        for (col_num, col_options) in col_options {
            // Check if the column number is contiguous with the previous column
            // and if the format is the same.
            if col_num == last_col + 1 && col_options == prev_col_options {
                last_col = col_num;
            } else {
                // If not write out the current range of columns and start again.
                self.write_col_element(first_col, last_col, &prev_col_options);
                first_col = col_num;
                last_col = first_col;
                prev_col_options = col_options;
            }
        }

        // We will exit the previous loop with one unhandled column range.
        self.write_col_element(first_col, last_col, &prev_col_options);

        xml_end_tag(&mut self.writer, "cols");
    }

    // Write the <col> element.
    fn write_col_element(&mut self, first_col: ColNum, last_col: ColNum, col_options: &ColOptions) {
        let first_col = first_col + 1;
        let last_col = last_col + 1;
        let mut width = col_options.width;
        let xf_index = col_options.xf_index;
        let has_custom_width = width != DEFAULT_COL_WIDTH;
        let hidden = col_options.hidden;

        // The default col width changes to 0 for hidden columns.
        if width == DEFAULT_COL_WIDTH && hidden {
            width = 0.0;
        }

        // Convert column width from user units to character width.
        if width > 0.0 {
            // Properties for Calibri 11.
            let max_digit_width = 7.0_f64;
            let padding = 5.0_f64;

            if width < 1.0 {
                width = ((width * (max_digit_width + padding)).round() / max_digit_width * 256.0)
                    .floor()
                    / 256.0;
            } else {
                width = (((width * max_digit_width).round() + padding) / max_digit_width * 256.0)
                    .floor()
                    / 256.0;
            }
        }

        let mut attributes = vec![
            ("min", first_col.to_string()),
            ("max", last_col.to_string()),
            ("width", width.to_string()),
        ];

        if xf_index > 0 {
            let xf_index = self.get_global_xf_index(xf_index);
            attributes.push(("style", xf_index.to_string()));
        }

        if col_options.autofit {
            attributes.push(("bestFit", "1".to_string()));
        }

        if hidden {
            attributes.push(("hidden", "1".to_string()));
        }

        if has_custom_width || hidden {
            attributes.push(("customWidth", "1".to_string()));
        }

        if col_options.level > 0 {
            attributes.push(("outlineLevel", col_options.level.to_string()));
        }

        if col_options.collapsed {
            attributes.push(("collapsed", "1".to_string()));
        }

        xml_empty_tag(&mut self.writer, "col", &attributes);
    }

    // Write the <headerFooter> element.
    fn write_header_footer(&mut self) {
        let mut attributes = vec![];

        if !self.header_footer_scale_with_doc {
            attributes.push(("scaleWithDoc", "0".to_string()));
        }

        if !self.header_footer_align_with_page {
            attributes.push(("alignWithMargins", "0".to_string()));
        }

        if self.header.is_empty() && self.footer.is_empty() {
            xml_empty_tag(&mut self.writer, "headerFooter", &attributes);
        } else {
            xml_start_tag(&mut self.writer, "headerFooter", &attributes);

            // Write the oddHeader element.
            if !self.header.is_empty() {
                self.write_odd_header();
            }

            // Write the oddFooter element.
            if !self.footer.is_empty() {
                self.write_odd_footer();
            }

            xml_end_tag(&mut self.writer, "headerFooter");
        }
    }

    // Write the <oddHeader> element.
    fn write_odd_header(&mut self) {
        let header = self
            .header
            .replace("&[Tab]", "&A")
            .replace("&[Date]", "&D")
            .replace("&[File]", "&F")
            .replace("&[Page]", "&P")
            .replace("&[Path]", "&Z")
            .replace("&[Time]", "&T")
            .replace("&[Pages]", "&N")
            .replace("&[Picture]", "&G");

        xml_data_element_only(&mut self.writer, "oddHeader", &header);
    }

    // Write the <oddFooter> element.
    fn write_odd_footer(&mut self) {
        let footer = self
            .footer
            .replace("&[Tab]", "&A")
            .replace("&[Date]", "&D")
            .replace("&[File]", "&F")
            .replace("&[Page]", "&P")
            .replace("&[Path]", "&Z")
            .replace("&[Time]", "&T")
            .replace("&[Pages]", "&N")
            .replace("&[Picture]", "&G");

        xml_data_element_only(&mut self.writer, "oddFooter", &footer);
    }

    // Write the <drawing> element.
    fn write_drawing(&mut self) {
        self.rel_count += 1;
        let attributes = [("r:id", format!("rId{}", self.rel_count))];

        xml_empty_tag(&mut self.writer, "drawing", &attributes);
    }

    // Write the <legacyDrawing> element.
    fn write_legacy_drawing(&mut self) {
        self.rel_count += 1;
        let attributes = [("r:id", format!("rId{}", self.rel_count))];

        xml_empty_tag(&mut self.writer, "legacyDrawing", &attributes);
    }

    // Write the <legacyDrawingHF> element.
    fn write_legacy_drawing_hf(&mut self) {
        self.rel_count += 1;
        let attributes = [("r:id", format!("rId{}", self.rel_count))];

        xml_empty_tag(&mut self.writer, "legacyDrawingHF", &attributes);
    }

    // Write the <picture> element.
    fn write_picture(&mut self) {
        self.rel_count += 1;
        let attributes = [("r:id", format!("rId{}", self.rel_count))];

        xml_empty_tag(&mut self.writer, "picture", &attributes);
    }

    // Write the <tableParts> element.
    fn write_table_parts(&mut self) {
        let num_tables = self.tables.len();

        let attributes = [("count", num_tables.to_string())];

        xml_start_tag(&mut self.writer, "tableParts", &attributes);

        for _ in 1..=num_tables {
            self.rel_count += 1;

            // Write the tablePart element.
            self.write_table_part(self.rel_count);
        }

        xml_end_tag(&mut self.writer, "tableParts");
    }

    // Write the <tablePart> element.
    fn write_table_part(&mut self, index: u32) {
        let attributes = [("r:id", format!("rId{index}"))];

        xml_empty_tag(&mut self.writer, "tablePart", &attributes);
    }

    // Write the <sheetProtection> element.
    fn write_sheet_protection(&mut self) {
        let mut attributes = vec![];

        if self.protection_hash != 0x0000 {
            attributes.push(("password", format!("{:04X}", self.protection_hash)));
        }

        if self.is_chartsheet {
            if self.protection_options.contents {
                attributes.push(("content", "1".to_string()));
            }

            if self.protection_options.edit_objects {
                attributes.push(("objects", "1".to_string()));
            }
        } else {
            attributes.push(("sheet", "1".to_string()));

            if !self.protection_options.edit_objects {
                attributes.push(("objects", "1".to_string()));
            }

            if !self.protection_options.edit_scenarios {
                attributes.push(("scenarios", "1".to_string()));
            }

            if self.protection_options.format_cells {
                attributes.push(("formatCells", "0".to_string()));
            }

            if self.protection_options.format_columns {
                attributes.push(("formatColumns", "0".to_string()));
            }

            if self.protection_options.format_rows {
                attributes.push(("formatRows", "0".to_string()));
            }

            if self.protection_options.insert_columns {
                attributes.push(("insertColumns", "0".to_string()));
            }

            if self.protection_options.insert_rows {
                attributes.push(("insertRows", "0".to_string()));
            }

            if self.protection_options.insert_links {
                attributes.push(("insertHyperlinks", "0".to_string()));
            }

            if self.protection_options.delete_columns {
                attributes.push(("deleteColumns", "0".to_string()));
            }

            if self.protection_options.delete_rows {
                attributes.push(("deleteRows", "0".to_string()));
            }

            if !self.protection_options.select_locked_cells {
                attributes.push(("selectLockedCells", "1".to_string()));
            }

            if self.protection_options.sort {
                attributes.push(("sort", "0".to_string()));
            }

            if self.protection_options.use_autofilter {
                attributes.push(("autoFilter", "0".to_string()));
            }

            if self.protection_options.use_pivot_tables {
                attributes.push(("pivotTables", "0".to_string()));
            }

            if !self.protection_options.select_unlocked_cells {
                attributes.push(("selectUnlockedCells", "1".to_string()));
            }
        }

        xml_empty_tag(&mut self.writer, "sheetProtection", &attributes);
    }

    // Write the <protectedRanges> element.
    fn write_protected_ranges(&mut self) {
        xml_start_tag_only(&mut self.writer, "protectedRanges");

        for (range, name, hash) in self.unprotected_ranges.clone() {
            self.write_protected_range(range, name, hash);
        }

        xml_end_tag(&mut self.writer, "protectedRanges");
    }

    // Write the <protectedRange> element.
    fn write_protected_range(&mut self, range: String, name: String, hash: u16) {
        let mut attributes = vec![];

        if hash > 0 {
            attributes.push(("password", format!("{hash:04X}")));
        }

        attributes.push(("sqref", range));
        attributes.push(("name", name));

        xml_empty_tag(&mut self.writer, "protectedRange", &attributes);
    }

    // Write the <rowBreaks> element.
    fn write_row_breaks(&mut self) {
        let attributes = [
            ("count", self.horizontal_breaks.len().to_string()),
            ("manualBreakCount", self.horizontal_breaks.len().to_string()),
        ];

        xml_start_tag(&mut self.writer, "rowBreaks", &attributes);

        for row_num in self.horizontal_breaks.clone() {
            // Write the brk element.
            self.write_row_brk(row_num);
        }

        xml_end_tag(&mut self.writer, "rowBreaks");
    }

    // Write the row <brk> element.
    fn write_row_brk(&mut self, row_num: u32) {
        let attributes = [
            ("id", row_num.to_string()),
            ("max", "16383".to_string()),
            ("man", "1".to_string()),
        ];

        xml_empty_tag(&mut self.writer, "brk", &attributes);
    }

    // Write the <colBreaks> element.
    fn write_col_breaks(&mut self) {
        let attributes = [
            ("count", self.vertical_breaks.len().to_string()),
            ("manualBreakCount", self.vertical_breaks.len().to_string()),
        ];

        xml_start_tag(&mut self.writer, "colBreaks", &attributes);

        for col_num in self.vertical_breaks.clone() {
            // Write the brk element.
            self.write_col_brk(col_num);
        }

        xml_end_tag(&mut self.writer, "colBreaks");
    }

    // Write the col <brk> element.
    fn write_col_brk(&mut self, col_num: u32) {
        let attributes = [
            ("id", col_num.to_string()),
            ("max", "1048575".to_string()),
            ("man", "1".to_string()),
        ];

        xml_empty_tag(&mut self.writer, "brk", &attributes);
    }

    // Write the <extLst> element.
    fn write_extensions(&mut self) {
        xml_start_tag_only(&mut self.writer, "extLst");

        // Write the x14:conditionalFormattings element.
        if self.has_x14_conditional_formats {
            let attributes = [
                (
                    "xmlns:x14",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                ),
                ("uri", "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"),
            ];
            xml_start_tag(&mut self.writer, "ext", &attributes);
            self.write_conditional_formattings();
        }

        // Write the x14:sparklineGroups element.
        if self.has_sparklines {
            let attributes = [
                (
                    "xmlns:x14",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                ),
                ("uri", "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"),
            ];
            xml_start_tag(&mut self.writer, "ext", &attributes);
            self.write_sparkline_groups();
        }

        xml_end_tag(&mut self.writer, "ext");

        xml_end_tag(&mut self.writer, "extLst");
    }

    // Write the <x14:sparklineGroups> element.
    fn write_sparkline_groups(&mut self) {
        let attributes = [(
            "xmlns:xm",
            "http://schemas.microsoft.com/office/excel/2006/main",
        )];

        xml_start_tag(&mut self.writer, "x14:sparklineGroups", &attributes);

        let sparklines = self.sparklines.clone();
        for sparkline in sparklines.into_iter().rev() {
            // Write the x14:sparklineGroup element.
            self.write_sparkline_group(&sparkline);
        }

        xml_end_tag(&mut self.writer, "x14:sparklineGroups");
    }

    // Write the <x14:sparklineGroup> element.
    fn write_sparkline_group(&mut self, sparkline: &Sparkline) {
        let mut attributes = vec![];

        if let Some(max) = sparkline.custom_max {
            attributes.push(("manualMax", max.to_string()));
        }

        if let Some(min) = sparkline.custom_min {
            attributes.push(("manualMin", min.to_string()));
        }

        match sparkline.sparkline_type {
            SparklineType::Column => {
                attributes.push(("type", "column".to_string()));
            }
            SparklineType::WinLose => {
                attributes.push(("type", "stacked".to_string()));
            }
            SparklineType::Line => {}
        }

        if let Some(weight) = sparkline.line_weight {
            attributes.push(("lineWeight", weight.to_string()));
        }

        if sparkline.date_range.has_data() {
            attributes.push(("dateAxis", "1".to_string()));
        }

        match sparkline.show_empty_cells_as {
            ChartEmptyCells::Gaps | ChartEmptyCells::Connected => {
                attributes.push((
                    "displayEmptyCellsAs",
                    sparkline.show_empty_cells_as.to_string(),
                ));
            }
            ChartEmptyCells::Zero => {}
        }

        if sparkline.show_markers {
            attributes.push(("markers", "1".to_string()));
        }

        if sparkline.show_high_point {
            attributes.push(("high", "1".to_string()));
        }

        if sparkline.show_low_point {
            attributes.push(("low", "1".to_string()));
        }

        if sparkline.show_first_point {
            attributes.push(("first", "1".to_string()));
        }

        if sparkline.show_last_point {
            attributes.push(("last", "1".to_string()));
        }

        if sparkline.show_negative_points {
            attributes.push(("negative", "1".to_string()));
        }

        if sparkline.show_axis {
            attributes.push(("displayXAxis", "1".to_string()));
        }

        if sparkline.show_hidden_data {
            attributes.push(("displayHidden", "1".to_string()));
        }

        if sparkline.custom_min.is_some() {
            attributes.push(("minAxisType", "custom".to_string()));
        } else if sparkline.group_min {
            attributes.push(("minAxisType", "group".to_string()));
        }

        if sparkline.custom_max.is_some() {
            attributes.push(("maxAxisType", "custom".to_string()));
        } else if sparkline.group_max {
            attributes.push(("maxAxisType", "group".to_string()));
        }

        if sparkline.show_right_to_left {
            attributes.push(("rightToLeft", "1".to_string()));
        }

        xml_start_tag(&mut self.writer, "x14:sparklineGroup", &attributes);

        // Write the sparkline color elements.
        self.write_sparkline_color("x14:colorSeries", sparkline.series_color);
        self.write_sparkline_color("x14:colorNegative", sparkline.negative_points_color);
        self.write_sparkline_color("x14:colorAxis", sparkline.axis_color);
        self.write_sparkline_color("x14:colorMarkers", sparkline.markers_color);
        self.write_sparkline_color("x14:colorFirst", sparkline.first_point_color);
        self.write_sparkline_color("x14:colorLast", sparkline.last_point_color);
        self.write_sparkline_color("x14:colorHigh", sparkline.high_point_color);
        self.write_sparkline_color("x14:colorLow", sparkline.low_point_color);

        if sparkline.date_range.has_data() {
            xml_data_element_only(&mut self.writer, "xm:f", &sparkline.date_range.formula());
        }

        // Write the x14:sparklines element.
        self.write_sparklines(sparkline);

        xml_end_tag(&mut self.writer, "x14:sparklineGroup");
    }

    // Write a sparkline <x14:color*> element.
    fn write_sparkline_color(&mut self, name: &str, color: Color) {
        xml_empty_tag(&mut self.writer, name, &color.attributes());
    }

    // Write the <x14:sparklines> element.
    fn write_sparklines(&mut self, sparkline: &Sparkline) {
        xml_start_tag_only(&mut self.writer, "x14:sparklines");

        for range in &sparkline.ranges {
            xml_start_tag_only(&mut self.writer, "x14:sparkline");

            xml_data_element_only(&mut self.writer, "xm:f", &range.1);
            xml_data_element_only(&mut self.writer, "xm:sqref", &range.0);

            xml_end_tag(&mut self.writer, "x14:sparkline");
        }
        xml_end_tag(&mut self.writer, "x14:sparklines");
    }

    // Write the <ignoredErrors> element.
    fn write_ignored_errors(&mut self) {
        xml_start_tag_only(&mut self.writer, "ignoredErrors");

        for error_type in IgnoreError::iterator() {
            let error_name = error_type.to_string();

            if let Some(error_range) = self.ignored_errors.get(&error_type) {
                let attributes = [
                    ("sqref", error_range.clone()),
                    (&error_name, "1".to_string()),
                ];
                xml_empty_tag(&mut self.writer, "ignoredError", &attributes);
            }
        }

        xml_end_tag(&mut self.writer, "ignoredErrors");
    }
}

// -----------------------------------------------------------------------
// Traits
// -----------------------------------------------------------------------

/// Trait to map user defined types to one of the supported Excel native types.
///
/// This trait allows you to map user defined types into a type that Excel
/// supports and to write it via [`Worksheet::write()`] and
/// [`Worksheet::write_with_format()`]. Both of these methods need to be
/// implemented for the trait. See the example below.
///
/// # Examples
///
/// Example of how to extend the the `rust_xlsxwriter` `write()` method using
/// the `IntoExcelData` trait to handle arbitrary user data that can be mapped
/// to one of the main Excel data types.
///
/// ```
/// # // This code is available in examples/app_write_generic_data.rs
/// #
/// use rust_xlsxwriter::*;
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Add a format for the dates.
///     let format = Format::new().set_num_format("yyyy-mm-dd");
///
///     // Make the first column wider for clarity.
///     worksheet.set_column_width(0, 12)?;
///
///     // Write user defined type instances that implement the IntoExcelData trait.
///     worksheet.write_with_format(0, 0, UnixTime::new(0), &format)?;
///     worksheet.write_with_format(1, 0, UnixTime::new(946598400), &format)?;
///     worksheet.write_with_format(2, 0, UnixTime::new(1672531200), &format)?;
///
///     // Save the file to disk.
///     workbook.save("write_generic.xlsx")?;
///
///     Ok(())
/// }
///
/// // For this example we create a simple struct type to represent a Unix time.
/// // This is the number of elapsed seconds since the epoch of January 1970 (UTC).
/// // See https://en.wikipedia.org/wiki/Unix_time. Note, this is for demonstration
/// // purposes only. The `ExcelDateTime` struct in `rust_xlsxwriter` can handle
/// // Unix timestamps.
/// pub struct UnixTime {
///     seconds: u64,
/// }
///
/// impl UnixTime {
///     pub fn new(seconds: u64) -> UnixTime {
///         UnixTime { seconds }
///     }
/// }
///
/// // Implement the IntoExcelData trait to map our new UnixTime struct into an
/// // Excel type.
/// //
/// // The relevant Excel type is f64 which is used to store dates and times (along
/// // with a number format). The Unix 1970 epoch equates to a date/number of
/// // 25569.0. For Unix times beyond that we divide by the number of seconds in the
/// // day (24 * 60 * 60) to get the Excel serial date.
/// //
/// // We need to implement two methods for the trait in order to write data with
/// // and without a format.
/// //
/// impl IntoExcelData for UnixTime {
///     fn write(
///         self,
///         worksheet: &mut Worksheet,
///         row: RowNum,
///         col: ColNum,
///     ) -> Result<&mut Worksheet, XlsxError> {
///         // Convert the Unix time to an Excel datetime.
///         let datetime = 25569.0 + (self.seconds as f64 / (24.0 * 60.0 * 60.0));
///
///         // Write the date as a number with a format.
///         worksheet.write_number(row, col, datetime)
///     }
///
///     fn write_with_format<'a>(
///         self,
///         worksheet: &'a mut Worksheet,
///         row: RowNum,
///         col: ColNum,
///         format: &Format,
///     ) -> Result<&'a mut Worksheet, XlsxError> {
///         // Convert the Unix time to an Excel datetime.
///         let datetime = 25569.0 + (self.seconds as f64 / (24.0 * 60.0 * 60.0));
///
///         // Write the date with the user supplied format.
///         worksheet.write_number_with_format(row, col, datetime, format)
///     }
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/write_generic.png">
///
pub trait IntoExcelData {
    /// Trait method to handle writing an unformatted type to Excel.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError>;

    /// Trait method to handle writing a formatted type to Excel.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError>;
}

macro_rules! write_string_trait_impl {
    ($($t:ty)*) => ($(
        impl IntoExcelData for $t {
            fn write(
                self,
                worksheet: &mut Worksheet,
                row: RowNum,
                col: ColNum,
            ) -> Result<&mut Worksheet, XlsxError> {
                worksheet.store_string(row, col, self.into(), None)
            }

            fn write_with_format<'a>(
                self,
                worksheet: &'a mut Worksheet,
                row: RowNum,
                col: ColNum,
                format: &Format,
            ) -> Result<&'a mut Worksheet, XlsxError> {
                worksheet.store_string(row, col, self.into(), Some(format))
            }
        }
    )*)
}
write_string_trait_impl!(&str &String String Cow<'_, str>);

macro_rules! write_number_trait_impl {
    ($($t:ty)*) => ($(
        impl IntoExcelData for $t {
            fn write(
                self,
                worksheet: &mut Worksheet,
                row: RowNum,
                col: ColNum,
            ) -> Result<&mut Worksheet, XlsxError> {
                worksheet.store_number(row, col, self, None)
            }

            fn write_with_format<'a>(
                self,
                worksheet: &'a mut Worksheet,
                row: RowNum,
                col: ColNum,
                format: &Format,
            ) -> Result<&'a mut Worksheet, XlsxError> {
                worksheet.store_number(row, col, self, Some(format))
            }
        }
    )*)
}
write_number_trait_impl!(u8 i8 u16 i16 u32 i32 f32 f64);

// Note: Excel doesn't support saving the full range of i64/u64 in f64 so this
// is documented as a loss of precision.
macro_rules! write_number_trait_impl {
    ($($t:ty)*) => ($(
        impl IntoExcelData for $t {
            fn write(
                self,
                worksheet: &mut Worksheet,
                row: RowNum,
                col: ColNum,
            ) -> Result<&mut Worksheet, XlsxError> {
                #[allow(clippy::cast_precision_loss)]
                worksheet.store_number(row, col, self as f64, None)
            }

            fn write_with_format<'a>(
                self,
                worksheet: &'a mut Worksheet,
                row: RowNum,
                col: ColNum,
                format: &Format,
            ) -> Result<&'a mut Worksheet, XlsxError> {
                #[allow(clippy::cast_precision_loss)]
                worksheet.store_number(row, col, self as f64, Some(format))
            }
        }
    )*)
}
write_number_trait_impl!(u64 i64);

impl IntoExcelData for bool {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        worksheet.store_boolean(row, col, self, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        worksheet.store_boolean(row, col, self, Some(format))
    }
}

impl IntoExcelData for &ExcelDateTime {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = self.to_excel();
        worksheet.store_datetime(row, col, number, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        let number = self.to_excel();
        worksheet.store_datetime(row, col, number, Some(format))
    }
}

impl IntoExcelData for ExcelDateTime {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = self.to_excel_serial_date();
        worksheet.store_datetime(row, col, number, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        let number = self.to_excel_serial_date();
        worksheet.store_datetime(row, col, number, Some(format))
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelData for &NaiveDateTime {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = ExcelDateTime::chrono_datetime_to_excel(self);
        worksheet.store_datetime(row, col, number, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        let number = ExcelDateTime::chrono_datetime_to_excel(self);
        worksheet.store_datetime(row, col, number, Some(format))
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelData for &NaiveDate {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = ExcelDateTime::chrono_date_to_excel(self);
        worksheet.store_datetime(row, col, number, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        let number = ExcelDateTime::chrono_date_to_excel(self);
        worksheet.store_datetime(row, col, number, Some(format))
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoExcelData for &NaiveTime {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        let number = ExcelDateTime::chrono_time_to_excel(self);
        worksheet.store_datetime(row, col, number, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        let number = ExcelDateTime::chrono_time_to_excel(self);
        worksheet.store_datetime(row, col, number, Some(format))
    }
}

#[cfg(feature = "rust_decimal")]
impl IntoExcelData for Decimal {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        let Some(number) = self.to_f64() else {
            return Err(XlsxError::ParameterError(format!(
                "Cannot represent Decimal {:?} as an Excel f64 value.",
                self
            )));
        };
        worksheet.store_number(row, col, number, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        let Some(number) = self.to_f64() else {
            return Err(XlsxError::ParameterError(format!(
                "Cannot represent Decimal {:?} as an Excel f64 value.",
                self
            )));
        };
        worksheet.store_number(row, col, number, Some(format))
    }
}

impl IntoExcelData for Formula {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        worksheet.store_formula(row, col, self, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        worksheet.store_formula(row, col, self, Some(format))
    }
}

impl IntoExcelData for &Formula {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        worksheet.store_formula(row, col, (*self).clone(), None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        worksheet.store_formula(row, col, (*self).clone(), Some(format))
    }
}

impl IntoExcelData for Url {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        worksheet.store_url(row, col, &self, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        worksheet.store_url(row, col, &self, Some(format))
    }
}

impl IntoExcelData for &Url {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        worksheet.store_url(row, col, self, None)
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        worksheet.store_url(row, col, self, Some(format))
    }
}

impl<T: IntoExcelData> IntoExcelData for Option<T> {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        match self {
            Some(data) => worksheet.write(row, col, data),
            None => Ok(worksheet),
        }
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        match self {
            Some(data) => worksheet.write_with_format(row, col, data, format),
            None => worksheet.write_blank(row, col, format),
        }
    }
}

impl<T: IntoExcelData, E: IntoExcelData> IntoExcelData for Result<T, E> {
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&mut Worksheet, XlsxError> {
        match self {
            Ok(data) => worksheet.write(row, col, data),
            Err(data) => worksheet.write(row, col, data),
        }
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        match self {
            Ok(data) => worksheet.write_with_format(row, col, data, format),
            Err(data) => worksheet.write_with_format(row, col, data, format),
        }
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------

// Round to the closest integer number of emu units.
fn round_to_emus(dimension: f64) -> f64 {
    (dimension * 9525.0).round()
}

// Struct to contain a cell range with some utility debug and other methods.
#[derive(Clone)]
pub(crate) struct CellRange {
    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) last_row: RowNum,
    pub(crate) last_col: ColNum,
}

impl CellRange {
    pub(crate) fn new(
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> CellRange {
        CellRange {
            first_row,
            first_col,
            last_row,
            last_col,
        }
    }

    pub(crate) fn to_range_string(&self) -> String {
        utility::cell_range(self.first_row, self.first_col, self.last_row, self.last_col)
    }

    pub(crate) fn to_error_string(&self) -> String {
        format!(
            "({}, {}, {}, {}) / {}",
            self.first_row,
            self.first_col,
            self.last_row,
            self.last_col,
            utility::cell_range(self.first_row, self.first_col, self.last_row, self.last_col)
        )
    }
}

impl Default for CellRange {
    fn default() -> Self {
        CellRange {
            first_row: ROW_MAX,
            first_col: COL_MAX,
            last_row: 0,
            last_col: 0,
        }
    }
}

#[derive(Clone)]
struct RowOptions {
    height: f64,
    xf_index: u32,
    level: u8,
    hidden: bool,
    collapsed: bool,
    format: Option<Format>,
}

#[derive(Clone, PartialEq)]
struct ColOptions {
    width: f64,
    xf_index: u32,
    level: u8,
    hidden: bool,
    collapsed: bool,
    autofit: bool,
    format: Option<Format>,
}

#[derive(Clone)]
enum CellType {
    ArrayFormula {
        formula: Box<str>,
        xf_index: u32,
        result: Box<str>,
        is_dynamic: bool,
        range: Box<str>,
    },
    Blank {
        xf_index: u32,
    },
    Boolean {
        boolean: bool,
        xf_index: u32,
    },
    Error {
        xf_index: u32,
        value: u32,
    },
    Formula {
        formula: Box<str>,
        xf_index: u32,
        result: Box<str>,
    },
    Number {
        number: f64,
        xf_index: u32,
    },
    DateTime {
        number: f64,
        xf_index: u32,
    },
    String {
        string: Arc<str>,
        xf_index: u32,
        string_id: Option<u32>,
    },
    RichString {
        string: Arc<str>,
        xf_index: u32,
        raw_string: Arc<str>,
        string_id: Option<u32>,
    },
    InlineString {
        string: Arc<str>,
        xf_index: u32,
    },
}

#[derive(Clone, Copy)]
enum PageView {
    Normal,
    PageLayout,
    PageBreaks,
}

#[derive(PartialEq)]
enum FilterType {
    Table,
    Worksheet,
}

#[derive(Clone)]
struct Panes {
    freeze_cell: (RowNum, ColNum),
    top_cell: (RowNum, ColNum),
}

impl Panes {
    fn is_empty(&self) -> bool {
        self.freeze_cell.0 == 0 && self.freeze_cell.1 == 0
    }

    fn top_left(&self) -> String {
        if self.top_cell.0 == 0 && self.top_cell.1 == 0 {
            utility::row_col_to_cell(self.freeze_cell.0, self.freeze_cell.1)
        } else {
            utility::row_col_to_cell(self.top_cell.0, self.top_cell.1)
        }
    }
}

// Struct to hold and transform data for the various defined names variants:
// user defined names, autofilters, print titles and print areas.
#[derive(Clone)]
pub(crate) struct DefinedName {
    pub(crate) in_use: bool,
    pub(crate) name: String,
    pub(crate) sort_name: String,
    pub(crate) range: String,
    pub(crate) quoted_sheet_name: String,
    pub(crate) index: u16,
    pub(crate) name_type: DefinedNameType,
    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) last_row: RowNum,
    pub(crate) last_col: ColNum,
}

impl DefinedName {
    pub(crate) fn new() -> DefinedName {
        DefinedName {
            in_use: false,
            name: String::new(),
            sort_name: String::new(),
            range: String::new(),
            quoted_sheet_name: String::new(),
            index: 0,
            name_type: DefinedNameType::Global,
            first_row: ROW_MAX,
            first_col: COL_MAX,
            last_row: 0,
            last_col: 0,
        }
    }

    pub(crate) fn initialize(&mut self, sheet_name: &str) {
        self.quoted_sheet_name = sheet_name.to_string();
        self.set_range();
        self.set_sort_name();
    }

    // Get the version of the defined name required by the App.xml file. Global
    // and Autofilter variants return the empty string and are ignored.
    pub(crate) fn app_name(&self) -> String {
        match self.name_type {
            DefinedNameType::Local => format!("{}!{}", self.quoted_sheet_name, self.name),
            DefinedNameType::PrintArea => format!("{}!Print_Area", self.quoted_sheet_name),
            DefinedNameType::Autofilter => String::new(),
            DefinedNameType::PrintTitles => format!("{}!Print_Titles", self.quoted_sheet_name),
            DefinedNameType::Global => {
                if self.range.contains('!') {
                    self.name.clone()
                } else {
                    String::new()
                }
            }
        }
    }

    pub(crate) fn name(&self) -> String {
        match self.name_type {
            DefinedNameType::PrintArea => "_xlnm.Print_Area".to_string(),
            DefinedNameType::Autofilter => "_xlnm._FilterDatabase".to_string(),
            DefinedNameType::PrintTitles => "_xlnm.Print_Titles".to_string(),
            _ => self.name.clone(),
        }
    }

    // The defined names are stored in a sorted order based on lowercase
    // and modified versions of the actual defined name.
    pub(crate) fn set_sort_name(&mut self) {
        let mut sort_name = match self.name_type {
            DefinedNameType::PrintArea => "Print_Area{}".to_string(),
            DefinedNameType::Autofilter => "_FilterDatabase{}".to_string(),
            DefinedNameType::PrintTitles => "Print_Titles".to_string(),
            _ => self.name.clone(),
        };

        sort_name = sort_name.replace('\'', "");
        self.sort_name = sort_name.to_lowercase();
    }

    pub(crate) fn set_range(&mut self) {
        match self.name_type {
            DefinedNameType::Autofilter | DefinedNameType::PrintArea => {
                let range;
                if self.first_col == 0 && self.last_col == COL_MAX - 1 {
                    // The print range is the entire column range, therefore we
                    // create a row only range.
                    range = format!("${}:${}", self.first_row + 1, self.last_row + 1);
                } else if self.first_row == 0 && self.last_row == ROW_MAX - 1 {
                    // The print range is the entire row range, therefore we
                    // create a column only range.
                    range = format!(
                        "${}:${}",
                        utility::column_number_to_name(self.first_col),
                        utility::column_number_to_name(self.last_col)
                    );
                } else {
                    // Otherwise handle it as a standard cell range.
                    range = utility::cell_range_absolute(
                        self.first_row,
                        self.first_col,
                        self.last_row,
                        self.last_col,
                    );
                }

                self.range = format!("{}!{}", self.quoted_sheet_name, range);
            }
            DefinedNameType::PrintTitles => {
                let mut range = String::new();

                if self.first_col != COL_MAX || self.last_col != 0 {
                    // Repeat columns.
                    range = format!(
                        "{}!${}:${}",
                        self.quoted_sheet_name,
                        utility::column_number_to_name(self.first_col),
                        utility::column_number_to_name(self.last_col)
                    );
                }

                if self.first_row != ROW_MAX || self.last_row != 0 {
                    // Repeat rows.
                    let row_range = format!(
                        "{}!${}:${}",
                        self.quoted_sheet_name,
                        self.first_row + 1,
                        self.last_row + 1
                    );

                    if range.is_empty() {
                        // The range is rows only.
                        range = row_range;
                    } else {
                        // Excel stores combined repeat rows and columns as a
                        // comma separated list.
                        range = format!("{range},{row_range}");
                    }
                }

                self.range = range;
            }
            _ => {}
        }
    }
}

/// The `IgnoreError` enum defines the Excel cell error types that can be
/// ignored.
///
/// The equivalent options in Excel are:
///
/// <img src="https://rustxlsxwriter.github.io/images/ignore_errors_dialog.png">
///
/// Note, some of the items shown in the above dialog such as "Misleading Number
/// Formats" aren't saved in the output file by Excel and can't be turned off
/// permanently.
///
/// The enum values are used with the [`Worksheet::ignore_error()`] and
/// [`Worksheet::ignore_error_range()`] methods.
///
#[derive(Clone, Debug, PartialEq, Eq, Hash, Copy)]
pub enum IgnoreError {
    /// Ignore errors/warnings for numbers stored as text.
    NumberStoredAsText,

    /// Ignore errors/warnings for formula evaluation errors (such as divide by
    /// zero).
    FormulaError,

    /// Ignore errors/warnings for formulas that differ from surrounding
    /// formulas.
    FormulaDiffers,

    /// Ignore errors/warnings for formulas that refer to empty cells.
    FormulaRefersToEmptyCells,

    /// Ignore errors/warnings for formulas that omit cells in a range.
    FormulaOmitsCells,

    /// Ignore errors/warnings for cells in a table that do not comply with
    /// applicable data validation rules.
    DataValidationError,

    /// Ignore errors/warnings for formulas that contain a two digit text
    /// representation of a year.
    TwoDigitTextYear,

    /// Ignore  errors/warnings for unlocked cells that contain formulas.
    UnlockedCellsWithFormula,

    /// Ignore errors/warnings for cell formulas that differ from the column
    /// formula.
    InconsistentColumnFormula,
}

impl IgnoreError {
    /// Simple iterator for `IgnoreError`.
    pub fn iterator() -> impl Iterator<Item = IgnoreError> {
        [
            Self::NumberStoredAsText,
            Self::FormulaError,
            Self::FormulaDiffers,
            Self::FormulaOmitsCells,
            Self::UnlockedCellsWithFormula,
            Self::InconsistentColumnFormula,
            Self::FormulaRefersToEmptyCells,
            Self::DataValidationError,
            Self::TwoDigitTextYear,
        ]
        .iter()
        .copied()
    }
}

impl fmt::Display for IgnoreError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::FormulaError => write!(f, "evalError"),
            Self::FormulaDiffers => write!(f, "formula"),
            Self::TwoDigitTextYear => write!(f, "TwoDigitTextYear"),
            Self::FormulaOmitsCells => write!(f, "formulaRange"),
            Self::DataValidationError => write!(f, "listDataValidation"),
            Self::NumberStoredAsText => write!(f, "numberStoredAsText"),
            Self::UnlockedCellsWithFormula => write!(f, "unlockedFormula"),
            Self::InconsistentColumnFormula => write!(f, "calculatedColumn"),
            Self::FormulaRefersToEmptyCells => write!(f, "emptyCellReference"),
        }
    }
}

#[derive(Clone, Debug)]
pub(crate) enum DefinedNameType {
    Autofilter,
    Global,
    Local,
    PrintArea,
    PrintTitles,
}

#[derive(Clone, Copy, Eq, PartialEq)]
pub(crate) enum Visible {
    Default,
    Hidden,
    VeryHidden,
}

pub(crate) enum BorderPosition {
    All,
    RowLeft,
    RowCenter,
    RowRight,
    ColTop,
    ColCenter,
    ColBottom,
    TopLeft,
    TopCenter,
    TopRight,
    CenterLeft,
    CenterCenter,
    CenterRight,
    BottomLeft,
    BottomCenter,
    BottomRight,
}
