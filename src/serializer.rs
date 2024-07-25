// serializer - A serde serializer for use with `rust_xlsxwriter`.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! # Working with Serde
//!
//! Serialization is the process of converting data from one format to another
//! using rules shared between the data and the serializer. The
//! [Serde](https://serde.rs) crate allow you to attach these type of rules to
//! data and it also provides traits that can be implemented by serializers to
//! convert that data into different formats. The `rust_xlsxwriter` crate
//! implements the Serde [`serde::Serializer`] trait for [`Worksheet`] structs
//! which allows you to serialize data directly to a worksheet.
//!
//! The following sections explains how to serialize Serde enabled data to an
//! Excel worksheet using `rust_xlsxwriter`.
//!
//!
//! Contents:
//!
//! - [How serialization works in
//!   `rust_xlsxwriter`](#how-serialization-works-in-rust_xlsxwriter)
//! - [Setting serialization headers](#setting-serialization-headers)
//! - [Renaming fields when serializing](#renaming-fields-when-serializing)
//! - [Skipping fields when serializing](#skipping-fields-when-serializing)
//! - [Setting serialization formatting](#setting-serialization-formatting)
//! - [Serializing dates and times](#serializing-dates-and-times)
//! - [Controlling Excel output via `XlsxSerialize` and struct
//!   attributes](#controlling-excel-output-via-xlsxserialize-and-struct-attributes)
//!   - [Container `xlsx` attributes](#container-xlsx-attributes)
//!   - [Field `xlsx` attributes](#field-xlsx-attributes)
//! - [Limitations of serializing to
//!   Excel](#limitations-of-serializing-to-excel)
//!
//! **Note**: This functionality requires the use of the `serde` feature flag
//! with `rust_xlsxwriter`:
//!
//! ```bash
//! cargo add rust_xlsxwriter -F serde
//! ```
//!
//!
//!
//!
//! ## How serialization works in `rust_xlsxwriter`
//!
//! Serialization with `rust_xlsxwriter` needs to take into consideration
//! that the target output is a 2D grid of cells into which the data can be
//! serialized. As such the focus is on serializing data types that map to this
//! 2D grid such as structs or compound collections of structs such as vectors
//! or tuples and it (currently) ignores compound types like maps.
//!
//! The image below shows the basic scheme for mapping a struct to a worksheet:
//! fields are mapped to a header and values are mapped to sequential cells
//! below the header.
//!
//! <img src="https://rustxlsxwriter.github.io/images/serialize_intro1.png">
//!
//! This scheme needs an initial (row, col) location from which to start
//! serializing to allow the data to be positioned anywhere on the worksheet.
//! Subsequent serializations will be in the same columns (for the target struct
//! type) but will be one row lower in the worksheet.
//!
//! The type name and fields of the struct being serialized is also required
//! information. We will look at that in more detail in the next section.
//!
//! Here is an example program that demonstrates the basic steps for serializing
//! data to an Excel worksheet:
//!
//! ```
//! # // This code is available in examples/doc_worksheet_serialize_intro2.rs
//! #
//! use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
//! use serde::{Serialize, Deserialize};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Add some formats to use with the serialization data.
//!     let header_format = Format::new()
//!         .set_bold()
//!         .set_border(FormatBorder::Thin)
//!         .set_background_color("C6E0B4");
//!
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     #[serde(rename_all = "PascalCase")]
//!     struct Student<'a> {
//!         name: &'a str,
//!         age: u8,
//!         id: u32,
//!     }
//!
//!     let students = [
//!         Student {
//!             name: "Aoife",
//!             age: 25,
//!             id: 564351,
//!         },
//!         Student {
//!             name: "Caoimhe",
//!             age: 21,
//!             id: 443287,
//!         },
//!     ];
//!
//!     // Set up the start location and headers of the data to be serialized.
//!     worksheet.deserialize_headers_with_format::<Student>(1, 3, &header_format)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&students)?;
//!
//!     // Save the file.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img src="https://rustxlsxwriter.github.io/images/serialize_intro2.png">
//!
//!
//!
//!
//! ## Setting serialization headers
//!
//! When serializing structs `rust_xlsxwriter` needs to know the location where
//! the serialization starts and also the type and field names of the struct
//! being serialized. The field names are used as headers and the type name
//! allows for several distinct structs to be serialized to the same worksheet.
//!
//! The worksheet methods that perform this function fall into three types:
//!
//! 1. For structs that only derive [`Serialize`] `rust_xlsxwriter` will use the
//!    same serialization it uses to write data to figure out the required
//!    fields and headers. This has the minor disadvantage that it needs an
//!    instance of the type (and not just the type itself) in order to work.
//!
//!    The following methods are available for this approach:
//!
//!    - [`Worksheet::serialize_headers()`]: Requires a concrete instance of the
//!      type of struct that you wish to serialize. The library uses this to
//!      infer the struct name and fields.
//!
//!    - [`Worksheet::serialize_headers_with_format()`]: This is similar to the
//!      previous method but it allows you to add a cell format for the headers.
//!
//!    - [`Worksheet::serialize_headers_with_options()`]: Similar to the
//!      previous methods but also allows configuration of the headers and
//!      fields via [`SerializeFieldOptions`].
//!
//! 2. For structs that derive [`Serialize`] and [`Deserialize`]
//!    `rust_xlsxwriter` can use deserialization to figure out the required
//!    fields and headers. This has the minor advantage that it only needs the
//!    type and doesn't require an instance of the type.
//!
//!    The following methods are available for this approach:
//!
//!    - [`Worksheet::deserialize_headers()`]: This only requires the type of
//!      struct that you wish to serialize and that it derives the
//!      [`Deserialize`] and [`Serialize`] traits. The library uses this to
//!      infer the struct name and fields (via deserialization).
//!
//!    - [`Worksheet::deserialize_headers_with_format()`]: This is similar to
//!      the previous method but it allows you to add a cell format for the
//!      headers.
//!
//!    - [`Worksheet::deserialize_headers_with_options()`]: Similar to the
//!      previous methods but also allows configuration of the headers and
//!      fields via [`SerializeFieldOptions`].
//!
//! 3. For structs that derive [`Serialize`] and [`XlsxSerialize`] (a
//!    `rust_xlsxwriter` provided trait) the library uses proc macros associated
//!    with `XlsxSerialize` to figure out the required fields and headers. This
//!    has the advantage of allowing the user to specify header and field
//!    formatting via struct attributes. See [Controlling Excel output via
//!    `XlsxSerialize` and struct
//!    attributes](#controlling-excel-output-via-xlsxserialize-and-struct-attributes)
//!    for details on how this works.
//!
//!    Since the formatting can be encoded in struct attributes this approach
//!    only requires one worksheet method:
//!
//!    - [`Worksheet::set_serialize_headers()`]: This only requires the type of
//!      struct that you wish to serialize and that it derives the
//!      [`XlsxSerialize`] and [`Serialize`] traits.
//!
//! The availability of 3 different approaches (and associated methods) will
//! probably cause mild anxiety in the reader as they will sense that they are
//! required to make a choice without knowing which is the best approach. In
//! general you can decide as follows: if your struct derives only have
//! [`Serialize`] and changing it is outside your control you must use the
//! serialization method (Approach 1); if you struct also supports
//! [`Deserialize`] you can save yourself the step of requiring a concrete
//! instance type and use the deserialization method (Approach 2); finally, if
//! the definition of the struct is under your control and you can add
//! [`XlsxSerialize`] then can define the formatting along with the struct and
//! also save yourself a lot of additional structural code to set up formatting
//! (Approach 3). The [`XlsxSerialize`] approach does, however,  limit you to
//! having a fixed set of formatting per struct.
//!
//! The two examples below should help inform your decision. If not then note
//! that the author mainly uses [`XlsxSerialize`] but you can use whichever
//! method works for you without fear of judgment.
//!
//! The first example below shows the usage of some of the serialization and
//! deserialization methods.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers3.rs
//! #
//! use rust_xlsxwriter::{
//!     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
//! };
//! use serde::{Deserialize, Serialize};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Set some column widths for clarity.
//!     worksheet.set_column_width(2, 4)?;
//!     worksheet.set_column_width(5, 4)?;
//!     worksheet.set_column_width(8, 4)?;
//!
//!     // Add some formats to use with the serialization data.
//!     let header_format = Format::new()
//!         .set_bold()
//!         .set_border(FormatBorder::Thin)
//!         .set_background_color("C6EFCE");
//!
//!     let currency_format = Format::new().set_num_format("$0.00");
//!
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let item1 = Produce {
//!         fruit: "Peach",
//!         cost: 1.05,
//!     };
//!
//!     let item2 = Produce {
//!         fruit: "Plum",
//!         cost: 0.15,
//!     };
//!
//!     let item3 = Produce {
//!         fruit: "Pear",
//!         cost: 0.75,
//!     };
//!
//!     // 1. Set the serialization location and headers with `deserialize_headers()`
//!     //    and serialize some data.
//!     worksheet.deserialize_headers::<Produce>(0, 0)?;
//!
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // 2. Set the serialization location and formatted headers with
//!     //    `serialize_headers_with_format()` and serialize some data.
//!     worksheet.serialize_headers_with_format(0, 3, &item1, &header_format)?;
//!
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // 3. Set the serialization location and headers with custom headers. We use
//!     //    the customization to set the header format and also rename the cell format
//!     //    for the number values.
//!     let custom_headers = [
//!         CustomSerializeField::new("fruit").rename("Item"),
//!         CustomSerializeField::new("cost")
//!             .rename("Price")
//!             .set_value_format(currency_format),
//!     ];
//!     let header_options = SerializeFieldOptions::new()
//!         .set_header_format(header_format)
//!         .set_custom_headers(&custom_headers);
//!
//!     // Set the serialization location and custom headers.
//!     worksheet.deserialize_headers_with_options::<Produce>(0, 6, &header_options)?;
//!
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // 4. Set the serialization location and headers with custom options. We use
//!     //    the customization to turn off the headers.
//!     let header_options = SerializeFieldOptions::new().hide_headers(true);
//!
//!     // Set the serialization location and custom headers.
//!     worksheet.serialize_headers_with_options(0, 9, &item1, &header_options)?;
//!
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // Save the file.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers3.png">
//!
//!
//! The second example uses `XlsxSerialize` and struct attributes to define the
//! output. Note that there is a lot less structural code in comparison to the
//! previous examples and that the formatting definitions are inline with the
//! struct. On the other hand you should also note that the previous example
//! where four different outputs are obtained for the same struct isn't possible
//! with this approach since the format definition is locked to the struct.
//!
//! ```
//! # // This code is available in examples/doc_xlsxserialize_intro.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
//! use serde::Serialize;
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Create a serializable struct.
//!     #[derive(XlsxSerialize, Serialize)]
//!     #[xlsx(header_format = Format::new().set_bold())]
//!     struct Produce {
//!         #[xlsx(rename = "Item")]
//!         #[xlsx(column_width = 12.0)]
//!         fruit: &'static str,
//!
//!         #[xlsx(rename = "Price", num_format = "$0.00")]
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let item1 = Produce {
//!         fruit: "Peach",
//!         cost: 1.05,
//!     };
//!
//!     let item2 = Produce {
//!         fruit: "Plum",
//!         cost: 0.15,
//!     };
//!
//!     let item3 = Produce {
//!         fruit: "Pear",
//!         cost: 0.75,
//!     };
//!
//!     // Set the serialization location and headers.
//!     worksheet.set_serialize_headers::<Produce>(0, 0)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // Save the file to disk.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_intro.png">
//!
//!
//!
//!
//! ## Renaming fields when serializing
//!
//! As explained above serialization converts the field names of structs to
//! column headers at the top of serialized data. The default field names are
//! generally lowercase and snake case and may not be the way you want the
//! header names displayed in Excel. In which case you can use one of the three
//! main methods to rename the fields/headers:
//!
//! 1. Rename the field during serialization using the Serde:
//!    - [field attribute]: `#[serde(rename = "name")` or
//!    - [container attribute]: `#[serde(rename_all = "...")]`.
//! 2. Rename the header (not field) when setting up custom serialization
//!    headers via [`Worksheet::deserialize_headers_with_options()`] or
//!    [`Worksheet::serialize_headers_with_options()`] and
//!    [`CustomSerializeField::rename()`].
//! 3. Rename the field using the [`XlsxSerialize`] field attribute
//!    `#[xlsx(rename = "name")`. This is effectively the same as the Serde
//!    attribute but allows you to specify a different name for the Excel
//!    serialization. See [Field `xlsx` attributes](#field-xlsx-attributes) for
//!    details.
//!
//! [field attribute]: https://serde.rs/field-attrs.html
//! [container attribute]: https://serde.rs/container-attrs.html
//!
//! Examples of these methods are shown below.
//!
//! ### Examples of field renaming
//!
//! The following example demonstrates renaming fields during serialization by
//! using Serde field attributes.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_rename1.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError};
//! # use serde::{Deserialize, Serialize};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Create a serializable struct. Note the serde attributes.
//!     #[derive(Deserialize, Serialize)]
//!     struct Produce {
//!         #[serde(rename = "Item")]
//!         fruit: &'static str,
//!
//!         #[serde(rename = "Price")]
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let items = [
//!         Produce {
//!             fruit: "Peach",
//!             cost: 1.05,
//!         },
//!         Produce {
//!             fruit: "Plum",
//!             cost: 0.15,
//!         },
//!         Produce {
//!             fruit: "Pear",
//!             cost: 0.75,
//!         },
//!     ];
//!
//!     // Set the serialization location and headers.
//!     worksheet.deserialize_headers::<Produce>(0, 0)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&items)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_rename1.png">
//!
//!
//! The following example demonstrates renaming fields during serialization by
//! specifying custom headers and renaming them there. The output is the same as
//! the image above.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_rename2.rs
//! #
//! # use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
//! # use serde::{Deserialize, Serialize};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let items = [
//!         Produce {
//!             fruit: "Peach",
//!             cost: 1.05,
//!         },
//!         Produce {
//!             fruit: "Plum",
//!             cost: 0.15,
//!         },
//!         Produce {
//!             fruit: "Pear",
//!             cost: 0.75,
//!         },
//!     ];
//!
//!     // Set up the custom headers.
//!     let custom_headers = [
//!         CustomSerializeField::new("fruit").rename("Item"),
//!         CustomSerializeField::new("cost").rename("Price"),
//!     ];
//!     let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);
//!
//!     // Set the serialization location and custom headers.
//!     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&items)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! The following example demonstrates renaming fields during serialization by
//! using `XlsxSerialize` `xlsx` field attributes. The output is the same as the
//! image above.
//!
//! ```rust
//! # // This code is available in examples/doc_xlsxserialize_rename.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
//! # use serde::Serialize;
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Create a serializable struct.
//!     #[derive(XlsxSerialize, Serialize)]
//!     struct Produce {
//!         #[xlsx(rename = "Item")]
//!         fruit: &'static str,
//!
//!         #[xlsx(rename = "Price")]
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let items = [
//!         Produce {
//!             fruit: "Peach",
//!             cost: 1.05,
//!         },
//!         Produce {
//!             fruit: "Plum",
//!             cost: 0.15,
//!         },
//!         Produce {
//!             fruit: "Pear",
//!             cost: 0.75,
//!         },
//!     ];
//!
//!     // Set the serialization location and headers.
//!     worksheet.set_serialize_headers::<Produce>(0, 0)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&items)?;
//! #
//! #     // Save the file to disk.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//!
//!
//!
//! ## Skipping fields when serializing
//!
//! When serializing a struct you may not want all of the fields to be
//! serialized. For example the struct may contain internal fields that aren't
//! of interest to the end user. There are several ways to skip fields:
//!
//! 1. Using the Serde [field attributes] `#[serde(skip)]`. This is the simplest
//!    and best method.
//! 2. Explicitly omitting the field by setting up custom serialization headers
//!    This method is useful when you can't add any additional attributes on the
//!    struct.
//! 3. Marking the field as skippable via custom headers and the
//!    [`CustomSerializeField::skip()`] method. This is only required in a few
//!    edge cases where other methods won't work.
//! 4. Skip the field using the [`XlsxSerialize`] field attribute
//!    `#[xlsx(skip)]`. This is effectively the same as the Serde attribute but
//!    allows you to specify a different behavior for the Excel serialization.
//!    See [Field `xlsx` attributes](#field-xlsx-attributes) for details.
//!
//! [field attributes]: https://serde.rs/field-attrs.html
//!
//! Examples of all four methods are shown below.
//!
//!
//! ### Examples of field skipping
//!
//! The following example demonstrates skipping fields during serialization by
//! using Serde field attributes. Since the field is no longer used we also need
//! to tell `rustc` not to emit a `dead_code` warning.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_skip1.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError};
//! use serde::{Deserialize, Serialize};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Create a serializable struct. Note the serde attribute.
//!     #[derive(Deserialize, Serialize)]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!
//!         #[serde(skip)]
//!         #[allow(dead_code)]
//!         in_stock: bool,
//!     }
//!
//!     // Create some data instances.
//!     let item1 = Produce {
//!         fruit: "Peach",
//!         cost: 1.05,
//!         in_stock: true,
//!     };
//!
//!     let item2 = Produce {
//!         fruit: "Plum",
//!         cost: 0.15,
//!         in_stock: true,
//!     };
//!
//!     let item3 = Produce {
//!         fruit: "Pear",
//!         cost: 0.75,
//!         in_stock: false,
//!     };
//!
//!     // Set the serialization location and headers.
//!     worksheet.deserialize_headers::<Produce>(0, 0)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // Save the file.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_skip1.png">
//!
//!
//!
//! The following example demonstrates skipping fields during serialization by
//! omitting them from the serialization headers. To do this we need to specify
//! custom headers and set [`SerializeFieldOptions::use_custom_headers_only()`].
//! The output is the same as the image above.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_skip2.rs
//! #
//! # use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
//! # use serde::{Deserialize, Serialize};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!         in_stock: bool,
//!     }
//!
//!     // Create some data instances.
//!     let item1 = Produce {
//!         fruit: "Peach",
//!         cost: 1.05,
//!         in_stock: true,
//!     };
//!
//!     let item2 = Produce {
//!         fruit: "Plum",
//!         cost: 0.15,
//!         in_stock: true,
//!     };
//!
//!     let item3 = Produce {
//!         fruit: "Pear",
//!         cost: 0.75,
//!         in_stock: false,
//!     };
//!
//!     // Set up only the custom headers we want and omit "in_stock".
//!     let custom_headers = [
//!         CustomSerializeField::new("fruit"),
//!         CustomSerializeField::new("cost"),
//!     ];
//!
//!     // Note the use of "use_custom_headers_only" to only serialize the named
//!     // custom headers.
//!     let header_options = SerializeFieldOptions::new()
//!         .use_custom_headers_only(true)
//!         .set_custom_headers(&custom_headers);
//!
//!     // Set the serialization location and custom headers.
//!     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! The following example is similar in setup to the previous example but
//! demonstrates skipping fields by explicitly skipping them in the custom
//! headers. This method should only be required in a few edge cases. The output
//! is the same as the image above.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_skip3.rs
//! #
//! # use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
//! # use serde::{Deserialize, Serialize};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//! #     // Create a serializable struct.
//! #     #[derive(Deserialize, Serialize)]
//! #     struct Produce {
//! #         fruit: &'static str,
//! #         cost: f64,
//! #         in_stock: bool,
//! #     }
//! #
//! #     // Create some data instances.
//! #     let item1 = Produce {
//! #         fruit: "Peach",
//! #         cost: 1.05,
//! #         in_stock: true,
//! #     };
//! #
//! #     let item2 = Produce {
//! #         fruit: "Plum",
//! #         cost: 0.15,
//! #         in_stock: true,
//! #     };
//! #
//! #     let item3 = Produce {
//! #         fruit: "Pear",
//! #         cost: 0.75,
//! #         in_stock: false,
//! #     };
//! #
//!     // We only need to set a custom header for the field we want to skip.
//!     let header_options = SerializeFieldOptions::new()
//!         .set_custom_headers(&[CustomSerializeField::new("in_stock").skip(true)]);
//!
//!     // Set the serialization location and custom headers.
//!     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
//! #
//! #     // Serialize the data.
//! #     worksheet.serialize(&item1)?;
//! #     worksheet.serialize(&item2)?;
//! #     worksheet.serialize(&item3)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! The following example demonstrates skipping fields during serialization by
//! using `XlsxSerialize` field attributes. Since the field is no longer used we
//! also need to tell `rustc` not to emit a `dead_code` warning.
//!
//! ```rust
//! # // This code is available in examples/doc_xlsxserialize_skip2.rs
//! #
//! # use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
//! # use serde::Serialize;
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     // Create a serializable struct.
//!     #[derive(XlsxSerialize, Serialize)]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!
//!         #[xlsx(skip)]
//!         #[allow(dead_code)]
//!         in_stock: bool,
//!     }
//! #
//! #     // Create some data instances.
//! #     let item1 = Produce {
//! #         fruit: "Peach",
//! #         cost: 1.05,
//! #         in_stock: true,
//! #     };
//! #
//! #     let item2 = Produce {
//! #         fruit: "Plum",
//! #         cost: 0.15,
//! #         in_stock: true,
//! #     };
//! #
//! #     let item3 = Produce {
//! #         fruit: "Pear",
//! #         cost: 0.75,
//! #         in_stock: false,
//! #     };
//! #
//! #     // Set the serialization location and headers.
//! #     worksheet.set_serialize_headers::<Produce>(0, 0)?;
//! #
//! #     // Serialize the data.
//! #     worksheet.serialize(&item1)?;
//! #     worksheet.serialize(&item2)?;
//! #     worksheet.serialize(&item3)?;
//! #
//! #     // Save the file to disk.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! ## Setting serialization formatting
//!
//! Serialization will transfer your data to a worksheet but it won't format it
//! without a few additional steps.
//!
//! The most common requirement is to format the header/fields at the top of the
//! serialized data. The simplest way to do this is to use the
//! [`Worksheet::deserialize_headers_with_format()`] or
//! [`Worksheet::serialize_headers_with_format()`] methods as shown in the
//! [Setting serialization headers](#setting-serialization-headers) section
//! above. Alternatively you can use the `#[xlsx(header_format = Format)]`
//! attribute (see [Container `xlsx` attributes](#container-xlsx-attributes)).
//!
//! The other common requirement is to format values that are serialized below
//! the headers such as numeric data where you need to control the number of
//! decimal places or make it appear as a currency.
//!
//! There are a few ways of formatting the values for a field:
//!
//! - Use [`Worksheet::set_column_format()`] to format the entire column.
//! - Use [`CustomSerializeField::set_column_format()`] to format the entire
//!   column. This is the same as the worksheet method but it has the advantage
//!   of having the column number calculated automatically based on the field
//!   name.
//! - Use [`CustomSerializeField::set_value_format()`] to format just the
//!   serialized data (and not the entire column).
//! - Using `XlsxSerialize` attributes. This is explained in detail in the
//!   section below on [Controlling Excel output via `XlsxSerialize` and struct
//!   attributes](#controlling-excel-output-via-xlsxserialize-and-struct-attributes).
//!
//! Examples of the first three methods are shown below. Examples of
//! `XlsxSerialize` attributes are shown in a subsequent section below.
//!
//! ### Examples of formatting
//!
//! The following example demonstrates serializing instances of a Serde derived
//! data structure to a worksheet with header and cell formatting.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_format4.rs
//! #
//! use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};
//! use serde::{Deserialize, Serialize};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Add some formats to use with the serialization data.
//!     let header_format = Format::new()
//!         .set_bold()
//!         .set_border(FormatBorder::Thin)
//!         .set_background_color("C6EFCE");
//!
//!     let currency_format = Format::new().set_num_format("$0.00");
//!
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     struct Produce {
//!         #[serde(rename = "Item")]
//!         fruit: &'static str,
//!
//!         #[serde(rename = "Price")]
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let item1 = Produce {
//!         fruit: "Peach",
//!         cost: 1.05,
//!     };
//!
//!     let item2 = Produce {
//!         fruit: "Plum",
//!         cost: 0.15,
//!     };
//!
//!     let item3 = Produce {
//!         fruit: "Pear",
//!         cost: 0.75,
//!     };
//!
//!     // Set a column format for the column. This is added to cell data that
//!     // doesn't have any other format so it doesn't affect the headers.
//!     worksheet.set_column_format(2, &currency_format)?;
//!
//!     // Set the serialization location and headers.
//!     worksheet.deserialize_headers_with_format::<Produce>(1, 1, &header_format)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // Save the file.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_custom.png">
//!
//! The following variation on the example demonstrates setting formatting via
//! custom headers. This produces the same output as the previous example but
//! doesn't require you to manually, or programmatically, calculate the column
//! number.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_format5.rs
//! #
//! # use rust_xlsxwriter::{
//! #     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
//! # };
//! # use serde::{Deserialize, Serialize};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//! #     // Add some formats to use with the serialization data.
//! #     let header_format = Format::new()
//! #         .set_bold()
//! #         .set_border(FormatBorder::Thin)
//! #         .set_background_color("C6EFCE");
//! #
//! #     let currency_format = Format::new().set_num_format("$0.00");
//! #
//! #     // Create a serializable struct.
//! #     #[derive(Deserialize, Serialize)]
//! #     struct Produce {
//! #         #[serde(rename = "Item")]
//! #         fruit: &'static str,
//! #
//! #         #[serde(rename = "Price")]
//! #         cost: f64,
//! #     }
//! #
//! #     // Create some data instances.
//! #     let item1 = Produce {
//! #         fruit: "Peach",
//! #         cost: 1.05,
//! #     };
//! #
//! #     let item2 = Produce {
//! #         fruit: "Plum",
//! #         cost: 0.15,
//! #     };
//! #
//! #     let item3 = Produce {
//! #         fruit: "Pear",
//! #         cost: 0.75,
//! #     };
//! #
//!     // Set up the custom headers.
//!     let custom_headers =
//!         [CustomSerializeField::new("Price").set_column_format(currency_format)];
//!
//!     let header_options = SerializeFieldOptions::new()
//!         .set_header_format(header_format)
//!         .set_custom_headers(&custom_headers);
//!
//!     // Set the serialization location and headers.
//!     worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;
//! #
//! #     // Serialize the data.
//! #     worksheet.serialize(&item1)?;
//! #     worksheet.serialize(&item2)?;
//! #     worksheet.serialize(&item3)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//!
//! The following variation on the example demonstrates setting formatting for
//! the serialized values, rather than the entire column. This allows a little
//! bit more precision on cell formatting.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_headers_format6.rs
//! #
//! # use rust_xlsxwriter::{
//! #     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
//! # };
//! # use serde::{Deserialize, Serialize};
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//! #     // Add some formats to use with the serialization data.
//! #     let header_format = Format::new()
//! #         .set_bold()
//! #         .set_border(FormatBorder::Thin)
//! #         .set_background_color("C6EFCE");
//! #
//! #     let currency_format = Format::new().set_num_format("$0.00");
//! #
//! #     // Create a serializable struct.
//! #     #[derive(Deserialize, Serialize)]
//! #     struct Produce {
//! #         #[serde(rename = "Item")]
//! #         fruit: &'static str,
//! #
//! #         #[serde(rename = "Price")]
//! #         cost: f64,
//! #     }
//! #
//! #     // Create some data instances.
//! #     let item1 = Produce {
//! #         fruit: "Peach",
//! #         cost: 1.05,
//! #     };
//! #
//! #     let item2 = Produce {
//! #         fruit: "Plum",
//! #         cost: 0.15,
//! #     };
//! #
//! #     let item3 = Produce {
//! #         fruit: "Pear",
//! #         cost: 0.75,
//! #     };
//! #
//!     // Set up the custom headers.
//!     let custom_headers =
//!         [CustomSerializeField::new("Price").set_value_format(currency_format)];
//!
//!     let header_options = SerializeFieldOptions::new()
//!         .set_header_format(header_format)
//!         .set_custom_headers(&custom_headers);
//!
//!     // Set the serialization location and headers.
//!     worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;
//! #
//! #     // Serialize the data.
//! #     worksheet.serialize(&item1)?;
//! #     worksheet.serialize(&item2)?;
//! #     worksheet.serialize(&item3)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//!
//!
//!
//! ## Serializing dates and times
//!
//! Dates and times can be serialized to Excel from one of the following types:
//!
//! - [`ExcelDateTime`]: The inbuilt `rust_xlsxwriter` datetime type.
//! - [`Chrono`] naive (i.e., timezone unaware) types:
//!   - [`NaiveDateTime`]
//!   - [`NaiveDate`]
//!   - [`NaiveTime`]
//!
//! [`ExcelDateTime`]: crate::ExcelDateTime
//! [`Chrono`]: https://docs.rs/chrono/latest/chrono
//! [`NaiveDate`]:
//!     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
//! [`NaiveTime`]:
//!     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
//! [`NaiveDateTime`]:
//!     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
//!
//! The [`ExcelDateTime`] type is serialized automatically since it implements
//! the [`Serialize`] trait. The [`Chrono`] types also implements [`Serialize`]
//! but they will serialize to an Excel string in RFC3339 format. To serialize
//! them to an Excel number/datetime format requires a serializing function like
//! [`Utility::serialize_chrono_naive_to_excel()`](crate::utility::serialize_chrono_naive_to_excel())
//! (as shown in the example below) or
//! [`Utility::serialize_chrono_option_naive_to_excel()`](crate::utility::serialize_chrono_option_naive_to_excel()).
//!
//! Excel datetimes also need a number format to display them as a date/time
//! since they are stored  as `f64` values. See [Datetimes in
//! Excel](crate::ExcelDateTime#datetimes-in-excel) and the previous section on
//! adding formatting.
//!
//! Note, Excel doesn't use timezones or try to convert or encode timezone
//! information in any way so they aren't supported by `rust_xlsxwriter`.
//!
//! ### Examples of serializing dates
//!
//! The following example demonstrates serializing instances of a Serde derived
//! data structure with [`ExcelDateTime`] fields.
//!
//! ```rust
//! # // This code is available in examples/doc_worksheet_serialize_datetime1.rs
//! #
//! use rust_xlsxwriter::{
//!     CustomSerializeField, ExcelDateTime, Format, FormatBorder,
//!     SerializeFieldOptions, Workbook, XlsxError,
//! };
//! use serde::{Deserialize, Serialize};
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Widen the date column for clarity.
//!     worksheet.set_column_width(1, 12)?;
//!
//!     // Add some formats to use with the serialization data.
//!     let header_format = Format::new()
//!         .set_bold()
//!         .set_border(FormatBorder::Thin)
//!         .set_background_color("C6E0B4");
//!
//!     let date_format = Format::new().set_num_format("yyyy/mm/dd");
//!
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     struct Student<'a> {
//!         name: &'a str,
//!         dob: ExcelDateTime,
//!         id: u32,
//!     }
//!
//!     let students = [
//!         Student {
//!             name: "Aoife",
//!             dob: ExcelDateTime::from_ymd(1998, 1, 12)?,
//!             id: 564351,
//!         },
//!         Student {
//!             name: "Caoimhe",
//!             dob: ExcelDateTime::from_ymd(2000, 5, 1)?,
//!             id: 443287,
//!         },
//!     ];
//!
//!     // Set up the start location and headers of the data to be serialized. Note,
//!     // we need to add a cell format for the datetime data.
//!     let custom_headers = [
//!         CustomSerializeField::new("name").rename("Student"),
//!         CustomSerializeField::new("dob")
//!             .rename("Birthday")
//!             .set_value_format(date_format),
//!         CustomSerializeField::new("id").rename("ID"),
//!     ];
//!     let header_options = SerializeFieldOptions::new()
//!         .set_header_format(header_format)
//!         .set_custom_headers(&custom_headers);
//!
//!     worksheet.deserialize_headers_with_options::<Student>(0, 0, &header_options)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&students)?;
//!
//!     // Save the file.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/worksheet_serialize_datetime1.png">
//!
//! Here is an example which serializes a struct with a [`NaiveDate`] field. The
//! output is the same as the previous example.
//!
//! ```ignore
//! # // This code is available in examples/doc_worksheet_serialize_datetime2.rs
//! #
//! use rust_xlsxwriter::utility::serialize_chrono_naive_to_excel;
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Widen the date column for clarity.
//!     worksheet.set_column_width(1, 12)?;
//!
//!     // Add some formats to use with the serialization data.
//!     let header_format = Format::new()
//!         .set_bold()
//!         .set_border(FormatBorder::Thin)
//!         .set_background_color("C6E0B4");
//!
//!     let date_format = Format::new().set_num_format("yyyy/mm/dd");
//!
//!     // Create a serializable struct.
//!     #[derive(Deserialize, Serialize)]
//!     struct Student<'a> {
//!         name: &'a str,
//!
//!         // Note, we add a `rust_xlsxwriter` function to serialize the date.
//!         #[serde(serialize_with = "serialize_chrono_naive_to_excel")]
//!         dob: NaiveDate,
//!
//!         id: u32,
//!     }
//! #
//! #     let students = [
//! #         Student {
//! #             name: "Aoife",
//! #             dob: NaiveDate::from_ymd_opt(1998, 1, 12).unwrap(),
//! #             id: 564351,
//! #         },
//! #         Student {
//! #             name: "Caoimhe",
//! #             dob: NaiveDate::from_ymd_opt(2000, 5, 1).unwrap(),
//! #             id: 443287,
//! #         },
//! #     ];
//! #
//! #     // Set up the start location and headers of the data to be serialized. Note,
//! #     // we need to add a cell format for the datetime data.
//! #     let custom_headers = [
//! #         CustomSerializeField::new("name").rename("Student"),
//! #         CustomSerializeField::new("dob")
//! #             .rename("Birthday")
//! #             .set_value_format(&date_format),
//! #         CustomSerializeField::new("id").rename("ID"),
//! #     ];
//! #     let header_options = SerializeFieldOptions::new()
//! #         .set_header_format(&header_format)
//! #         .set_custom_headers(&custom_headers);
//! #
//! #     worksheet.deserialize_headers_with_options::<Student>(0, 0, &header_options)?;
//! #
//! #     // Serialize the data.
//! #     worksheet.serialize(&students)?;
//! #
//! #     // Save the file.
//! #     workbook.save("serialize.xlsx")?;
//!
//!     // ...
//!
//!     Ok(())
//! }
//! ```
//!
//!
//!
//!
//!
//!
//! ## Controlling Excel output via `XlsxSerialize` and struct attributes
//!
//! In the sections above we saw how to use [Serde
//! Attributes](https://serde.rs/attributes.html) to control some aspects for
//! the serialization output.
//!
//! It is also possible to use the `XlsxSerialize` derive trait and associate
//! attributes to set `rust_xlsxwriter` specific formatting and options. For
//! example:
//!
//!
//! ```
//! # // This code is available in examples/doc_xlsxserialize_intro.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
//! use serde::Serialize;
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Create a serializable struct.
//!     #[derive(XlsxSerialize, Serialize)]
//!     #[xlsx(header_format = Format::new().set_bold())]
//!     struct Produce {
//!         #[xlsx(rename = "Item")]
//!         #[xlsx(column_width = 12.0)]
//!         fruit: &'static str,
//!
//!         #[xlsx(rename = "Price", num_format = "$0.00")]
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let item1 = Produce {
//!         fruit: "Peach",
//!         cost: 1.05,
//!     };
//!
//!     let item2 = Produce {
//!         fruit: "Plum",
//!         cost: 0.15,
//!     };
//!
//!     let item3 = Produce {
//!         fruit: "Pear",
//!         cost: 0.75,
//!     };
//!
//!     // Set the serialization location and headers.
//!     worksheet.set_serialize_headers::<Produce>(0, 0)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&item1)?;
//!     worksheet.serialize(&item2)?;
//!     worksheet.serialize(&item3)?;
//!
//!     // Save the file to disk.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Note the change of column width in Column A, the renamed headers and the
//! currency format in Column B numbers.
//!
//! <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_intro.png">
//!
//! `XlsxSerializer` provides the same formatting functionality as
//! [`SerializeFieldOptions`] and [`CustomSerializeField`] (in fact it used them
//! in the background). The advantage of using `XlsxSerializer` is that the
//! formatting output can be kept with the data format and seen in the same
//! context. Some users may find this to be a cleaner abstraction.
//!
//!
//!
//!
//! ## The `XlsxSerializer` attributes
//!
//! The `XlsxSerializer` derived trait adds a wrapper around a Serde
//! serializable struct to add `rust_xlsxwriter` specific formatting options. It
//! achieves this via the `rust_xlsxwriter` [`SerializeFieldOptions`] and
//! [`CustomSerializeField`] serialization configuration structs. It is
//! available when you enable the `serde` feature.
//!
//! As in Serde the [attributes](https://serde.rs/attributes.html) are divided
//! into "Container Attributes" which apply to the entire struct and "Field
//! Attributes" which apply to individuals fields. In an Excel context the
//! container attributes apply to all headers and field attributes apply to the
//! individual headers and the data below them.
//!
//! In order to demonstrate each attribute and its effect we will use variations
//! of the following example with the relevant attribute applied in the next
//! sections.
//!
//! ### Base example (without attributes applied)
//!
//! ```
//! # // This code is available in examples/doc_xlsxserialize_base.rs
//! #
//! use rust_xlsxwriter::{Workbook, XlsxError, XlsxSerialize};
//! use serde::Serialize;
//!
//! fn main() -> Result<(), XlsxError> {
//!     let mut workbook = Workbook::new();
//!
//!     // Add a worksheet to the workbook.
//!     let worksheet = workbook.add_worksheet();
//!
//!     // Create a serializable struct.
//!     #[derive(XlsxSerialize, Serialize)]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!     }
//!
//!     // Create some data instances.
//!     let items = [
//!         Produce {
//!             fruit: "Peach",
//!             cost: 1.05,
//!         },
//!         Produce {
//!             fruit: "Plum",
//!             cost: 0.15,
//!         },
//!         Produce {
//!             fruit: "Pear",
//!             cost: 0.75,
//!         },
//!     ];
//!
//!     // Set the serialization location and headers.
//!     worksheet.set_serialize_headers::<Produce>(0, 0)?;
//!
//!     // Serialize the data.
//!     worksheet.serialize(&items)?;
//!
//!     // Save the file to disk.
//!     workbook.save("serialize.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! ### Container `xlsx` attributes
//!
//! The following are the "Container" attributes supported by `XlsxSerializer`:
//!
//! - `#[xlsx(header_format = Format)`
//!
//!   The `header_format` container attribute sets the [`Format`] property for
//!   headers. See the [Working with attribute
//!   Formats](#working-with-attribute-formats) section below for information on
//!   handling formats.
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         #[xlsx(header_format = Format::new()
//!                .set_bold()
//!                .set_border(FormatBorder::Thin)
//!                .set_background_color("C6EFCE"))]
//!         struct Produce {
//!             fruit: &'static str,
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//! <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_header_format.png">
//!
//!
//! - `#[xlsx(hide_headers)`
//!
//!   The `hide_headers` container attribute hides the serialization headers:
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         #[xlsx(hide_headers)]
//!         struct Produce {
//!             fruit: &'static str,
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_hide_headers.png">
//!
//!
//! - `#[xlsx(table_default)`
//!
//!   The `table_default` container attribute adds a worksheet [`Table`]
//!   structure with default formatting to the serialized area:
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         #[xlsx(table_default)]
//!         struct Produce {
//!             fruit: &'static str,
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_table_default.png">
//!
//!
//! - `#[xlsx(table_style)`
//!
//!   The `table_style` container attribute adds a worksheet [`Table`]
//!   structure with a user specified [`TableStyle`] to the serialized area:
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         #[xlsx(table_style = TableStyle::Medium10)]
//!         struct Produce {
//!             fruit: &'static str,
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_table_style.png">
//!
//!
//! - `#[xlsx(table = Table)`
//!
//!   The `table` container attribute adds a user defined worksheet [`Table`]
//!   structure to the serialized area:
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         #[xlsx(table = Table::new())]
//!         struct Produce {
//!             fruit: &'static str,
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_table_default.png">
//!
//!   See the [Working with attribute
//!   Formats](#working-with-attribute-formats) section below for information on
//!   how to wrap complex objects like [`Format`] or [`Table`] in a function so
//!   it can be used as an attribute parameter.
//!
//!
//! ### Field `xlsx` attributes
//!
//! The following are the "Field" attributes supported by `XlsxSerializer`:
//!
//! - `#[xlsx(rename = "")`
//!
//!   The `rename` field attribute renames the Excel header. This is similar to
//!   the `#[serde(rename = "")` field attribute except that it doesn't rename
//!   the field for other types of serialization.
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         struct Produce {
//!             #[xlsx(rename = "Item")]
//!             fruit: &'static str,
//!
//!             #[xlsx(rename = "Price")]
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_rename.png">
//!
//!
//! - `#[xlsx(num_format = "")`
//!
//!   The `num_format` field attribute sets the property to change the number
//!   formatting of the output. It is a syntactic shortcut for
//!   [`Format::set_num_format()`].
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         struct Produce {
//!             fruit: &'static str,
//!
//!             #[xlsx(num_format = "$0.00")]
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_num_format.png">
//!
//!
//! - `#[xlsx(value_format = Format)`
//!
//!   The `value_format` field attribute sets the [`Format`] property for the
//!   fields below the header.
//!
//!   See the [Working with attribute Formats](#working-with-attribute-formats)
//!   section below for information on handling formats.
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         struct Produce {
//!             #[xlsx(value_format = Format::new().set_font_color("#FF0000"))]
//!             fruit: &'static str,
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_value_format.png">
//!
//!
//! - `#[xlsx(column_format = Format)`
//!
//!   The `column_format` field attribute is similar to the previous
//!   `value_format` attribute except that it sets the format for the entire
//!   column, including any data added manually below the serialized data.
//!
//!   <br>
//!
//! - `#[xlsx(header_format = Format)`
//!
//!   The `header_format` field attribute sets the [`Format`] property for the
//!   the header. It is similar to the container method of the same name except
//!   it only sets the format or one header:
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         struct Produce {
//!             fruit: &'static str,
//!
//!             #[xlsx(header_format = Format::new().set_bold().set_font_color("#FF0000"))]
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_field_header_format.png">
//!
//!
//! - `#[xlsx(column_width = float)`
//!
//!   The `column_width` field attribute sets the column width in character
//!   units.
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         struct Produce {
//!             fruit: &'static str,
//!             #[xlsx(column_width = 20.0)]
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img
//!   src="https://rustxlsxwriter.github.io/images/xlsxserialize_column_width.png">
//!
//!
//!
//! - `#[xlsx(column_width_pixels = int)`
//!
//!   The `column_width_pixels` field attribute field attribute is similar to
//!   the previous `column_width` attribute except that the width is specified
//!   in integer pixel units.
//!
//!   <br>
//!
//!
//!
//!
//! - `#[xlsx(skip)`
//!
//!   The `skip` field attribute skips writing the field to the target Excel
//!   file.  This is similar to the `#[serde(skip)` field attribute except that
//!   it doesn't skip the field for other types of serialization.
//!
//!   ```
//!   # use rust_xlsxwriter::XlsxSerialize;
//!   # use serde::Serialize;
//!   #
//!   # fn main() {
//!         #[derive(XlsxSerialize, Serialize)]
//!         struct Produce {
//!             fruit: &'static str,
//!
//!             #[xlsx(skip)]
//!             cost: f64,
//!         }
//!   # }
//!   ```
//!
//!   <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_skip.png">
//!
//!
//! Note, if required you can group more than one attribute
//!
//! ```text
//! #[xlsx(rename = "Item", column_width = 20.0)]
//! ```
//!
//! ### Working with attribute Formats
//!
//! When working with `XlsxSerialize` attributes that deal with [`Format`]
//! objects the definition can become quite long for an attribute value:
//!
//! ```
//! # use rust_xlsxwriter::XlsxSerialize;
//! # use serde::Serialize;
//! #
//! # fn main() {
//!       #[derive(XlsxSerialize, Serialize)]
//!       #[xlsx(header_format = Format::new()
//!              .set_bold()
//!              .set_border(FormatBorder::Thin)
//!              .set_background_color("C6EFCE"))]
//!       struct Produce {
//!           fruit: &'static str,
//!           cost: f64,
//!       }
//! # }
//! ```
//!
//! <br>
//!
//! You might in this case be tempted to define the format in another part of
//! your code and use a variable to define the format:
//!
//! ```compile_fail
//! # use rust_xlsxwriter::{Format, FormatBorder, XlsxSerialize};
//! # use serde::Serialize;
//! #
//! # fn main() {
//!     let my_header_format = Format::new()
//!         .set_bold()
//!         .set_border(FormatBorder::Thin)
//!         .set_background_color("C6EFCE");
//!
//!     #[derive(XlsxSerialize, Serialize)]
//!     #[xlsx(header_format = &my_header_format)] // Won't work.
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!     }
//! # }
//! ```
//!
//! <br>
//!
//! However, this won't work because Rust derive/proc macros are compiled
//! statically and `&header_format` is a dynamic variable.
//!
//! A workaround for this it to define any formats you wish to use in functions:
//!
//! ```
//! # // This code is available in examples/doc_xlsxserialize_header_format_reuse.rs
//! #
//! # use rust_xlsxwriter::{FormatBorder, Workbook, XlsxError, XlsxSerialize, Format};
//! # use serde::Serialize;
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     let mut workbook = Workbook::new();
//! #
//! #     // Add a worksheet to the workbook.
//! #     let worksheet = workbook.add_worksheet();
//! #
//!     fn my_header_format() -> Format {
//!         Format::new()
//!             .set_bold()
//!             .set_border(FormatBorder::Thin)
//!             .set_background_color("C6EFCE")
//!     }
//!
//!     #[derive(XlsxSerialize, Serialize)]
//!     #[xlsx(header_format = my_header_format())]
//!     struct Produce {
//!         fruit: &'static str,
//!         cost: f64,
//!     }
//! #
//! #     // Create some data instances.
//! #     let items = [
//! #         Produce {
//! #             fruit: "Peach",
//! #             cost: 1.05,
//! #         },
//! #         Produce {
//! #             fruit: "Plum",
//! #             cost: 0.15,
//! #         },
//! #         Produce {
//! #             fruit: "Pear",
//! #             cost: 0.75,
//! #         },
//! #     ];
//! #
//! #     // Set the serialization location and headers.
//! #     worksheet.set_serialize_headers::<Produce>(0, 0)?;
//! #
//! #     // Serialize the data.
//! #     worksheet.serialize(&items)?;
//! #
//! #     // Save the file to disk.
//! #     workbook.save("serialize.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/xlsxserialize_header_format_reuse.png">
//!
//!
//!
//!
//! ## Limitations of serializing to Excel
//!
//! The cell/grid format of Excel sets a physical limitation on what can be
//! serialized to a worksheet. Unlike other formats such as JSON or XML you
//! cannot serialize arbitrary nested data to Excel without making some
//! concessions to either the format or the contents of the data. When
//! serializing data to Excel via `rust_xlsxwriter` it is best to consider what
//! that data will look like while designing your serialization.
//!
//! Another limitation is that currently you can only serialize structs or
//! struct values in compound containers such as vectors. Not all of the
//! supported types in the [Serde data model] make sense in the context of
//! Excel. In upcoming releases I will try to add support for additional types
//! where it makes sense. If you have a valid use case please open a GitHub
//! issue to discuss it with an example data structure. Also Serde sub-struct
//! flattening is not currently supported.
//!
//! [Serde data model]: https://serde.rs/data-model.html
//!
//! Finally if you hit some serialization limitation using `rust_xlsxwriter`
//! remember that there are other non-serialization options available to use in
//! the standard [`Worksheet`] API to write scalar, vector and matrix data
//! types:
//!
//! - [`Worksheet::write()`]
//! - [`Worksheet::write_row()`]
//! - [`Worksheet::write_column()`]
//! - [`Worksheet::write_row_matrix()`]
//! - [`Worksheet::write_column_matrix()`]
//!
//! Magic is great but the direct approach will also work. Remember Terry
//! Pratchett's witches.
//!
#![warn(missing_docs)]

use std::collections::HashMap;
use std::sync::Arc;

use crate::{ColNum, Format, RowNum, Table, TableStyle, Worksheet, XlsxError};
use serde::de::Visitor;
use serde::{ser, Deserialize, Deserializer, Serialize};

// Convenience tuple struct Table data used for serialization formatting.
pub(crate) struct TableData(
    pub(crate) RowNum,
    pub(crate) ColNum,
    pub(crate) RowNum,
    pub(crate) ColNum,
    pub(crate) Table,
);

// -----------------------------------------------------------------------
// SerializerState, a struct to maintain row/column state and other metadata
// between serialized writes. This avoids passing around cell location
// information in the serializer.
// -----------------------------------------------------------------------
pub(crate) struct SerializerState {
    pub(crate) structs: HashMap<String, SerializationHeaderConfig>,
    pub(crate) current_struct: String,
    pub(crate) current_field: String,
}

impl SerializerState {
    // Create a new SerializerState struct.
    pub(crate) fn new() -> SerializerState {
        SerializerState {
            structs: HashMap::new(),
            current_struct: String::new(),
            current_field: String::new(),
        }
    }

    // Check if the current struct/field have been selected to be serialized by
    // the user. If it has then return the row value for the next `write()` call.
    pub(crate) fn current_state(&mut self) -> Result<(RowNum, ColNum, Arc<Option<Format>>), ()> {
        let Some(header_config) = self.structs.get_mut(&self.current_struct) else {
            return Err(());
        };

        let Some(field) = header_config.fields.get_mut(&self.current_field) else {
            return Err(());
        };

        // Set the "current" cell values used to write the serialized data.
        let row = header_config.max_row - 1;
        let col = field.col;
        let value_format = Arc::clone(&field.value_format);

        Ok((row, col, value_format))
    }

    // Store the name and max row of the current struct being serialized.
    pub(crate) fn set_current_struct(&mut self, struct_name: &str) {
        if struct_name != self.current_struct {
            self.current_struct = struct_name.to_string();
        }

        // Increment the max row every time we serialize a new struct instance.
        let Some(header_config) = self.structs.get_mut(&self.current_struct) else {
            return;
        };
        header_config.max_row += 1;
    }

    // Get dimensions of a serialization area. This is the internal function for
    // worksheet.get_serialize_dimensions().
    pub(crate) fn get_dimensions(
        &mut self,
        name: &str,
    ) -> Result<(RowNum, ColNum, RowNum, ColNum), XlsxError> {
        let Some(header_config) = self.structs.get(name) else {
            return Err(XlsxError::ParameterError(format!(
                "Unknown serialized struct '{name}'"
            )));
        };

        Ok((
            header_config.min_row,
            header_config.min_col,
            header_config.max_row - 1,
            header_config.max_col,
        ))
    }

    // Get dimensions of a column in a serialization area. This is the internal
    // function for worksheet.get_serialize_column_dimensions().
    pub(crate) fn get_column_dimensions(
        &mut self,
        struct_name: &str,
        field_name: &str,
    ) -> Result<(RowNum, ColNum, RowNum, ColNum), XlsxError> {
        let Some(header_config) = self.structs.get(struct_name) else {
            return Err(XlsxError::ParameterError(format!(
                "Unknown serialized struct '{struct_name}'"
            )));
        };

        let Some(field) = header_config.fields.get(field_name) else {
            return Err(XlsxError::ParameterError(format!(
                "Unknown serialized field '{field_name}'"
            )));
        };

        Ok((
            header_config.min_row,
            field.col,
            header_config.max_row - 1,
            field.col,
        ))
    }

    // Get all/any tables defined for serialization areas.
    pub(crate) fn get_tables(&mut self) -> Vec<TableData> {
        let mut tables = vec![];

        for header_config in self.structs.values_mut() {
            if let Some(table) = header_config.get_table() {
                tables.push(table);
            }
        }

        tables
    }
}

// -----------------------------------------------------------------------
// HeaderConfig, a struct to capture the metadata for fields associated
// with a struct.
// -----------------------------------------------------------------------
pub(crate) struct SerializationHeaderConfig {
    pub(crate) fields: HashMap<String, CustomSerializeField>,
    pub(crate) min_row: RowNum,
    pub(crate) min_col: ColNum,
    pub(crate) max_row: RowNum,
    pub(crate) max_col: ColNum,
    pub(crate) table: Option<Table>,
}

impl SerializationHeaderConfig {
    // Get table object and dimensions for the a serialization area.
    pub(crate) fn get_table(&mut self) -> Option<TableData> {
        let table = self.table.take();

        match table {
            Some(table) => Some(TableData(
                self.min_row,
                self.min_col,
                self.max_row - 1,
                self.max_col,
                table,
            )),
            None => None,
        }
    }
}

// -----------------------------------------------------------------------
// SerializeFieldOptions.
// -----------------------------------------------------------------------

/// The `SerializeFieldOptions` struct represents custom field/header options.
///
/// `SerializeFieldOptions` can be used to set column headers to map serialized
/// data to. It allows you to reorder, rename, format or skip headers and also
/// define formatting for field values.
///
/// It is used in conjunction with the [`CustomSerializeField`] struct and
/// [`Worksheet::deserialize_headers_with_options()`] and
/// [`Worksheet::serialize_headers_with_options()`] methods.
///
/// See [Working with Serde](crate::serializer) for an introduction to
/// serialization with `rust_xlsxwriter`.
///
/// # Examples
///
/// The following example demonstrates serializing instances of a Serde derived
/// data structure to a worksheet with custom headers and cell formatting.
///
/// ```
/// # // This code is available in examples/doc_worksheet_serialize_headers_custom.rs
/// #
/// # use rust_xlsxwriter::{
/// #     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
/// # };
/// # use serde::{Deserialize, Serialize};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some formats to use with the serialization data.
/// #     let header_format = Format::new()
/// #         .set_bold()
/// #         .set_border(FormatBorder::Thin)
/// #         .set_background_color("C6EFCE");
/// #
/// #     let currency_format = Format::new().set_num_format("$0.00");
/// #
/// #     // Create a serializable struct.
///     #[derive(Deserialize, Serialize)]
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
///     // Set up the custom headers.
///     let custom_headers = [
///         CustomSerializeField::new("fruit")
///             .rename("Item"),
///         CustomSerializeField::new("cost")
///             .rename("Price")
///             .set_value_format(currency_format),
///     ];
///
///     let header_options = SerializeFieldOptions::new()
///         .set_header_format(header_format)
///         .set_custom_headers(&custom_headers);
///
///     // Set the serialization location and custom headers.
///     worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;
///
///     // Serialize the data.
///     worksheet.serialize(&item1)?;
///     worksheet.serialize(&item2)?;
///     worksheet.serialize(&item3)?;
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
/// <img src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_custom.png">
///
#[derive(Clone)]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub struct SerializeFieldOptions {
    pub(crate) struct_name: String,
    pub(crate) header_format: Option<Format>,
    pub(crate) has_headers: bool,
    pub(crate) custom_headers: Vec<CustomSerializeField>,
    pub(crate) use_custom_headers_only: bool,
    pub(crate) table: Option<Table>,
}

impl Default for SerializeFieldOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl SerializeFieldOptions {
    /// Create serialization header options.
    ///
    /// Create a `SerializeFieldOptions` struct to be used with the
    /// [`Worksheet::deserialize_headers_with_options()`] or
    /// [`Worksheet::serialize_headers_with_options()`] methods.
    ///
    /// This can be used to set or skip column headers as well as specifying
    /// column and field specific formatting.
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn new() -> SerializeFieldOptions {
        SerializeFieldOptions {
            struct_name: String::new(),
            header_format: None,
            has_headers: true,
            custom_headers: vec![],
            use_custom_headers_only: false,
            table: None,
        }
    }

    /// Set the header format for a serialization headers.
    ///
    /// See [`Format`] for more information on formatting.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`Format`] property for the header.
    ///
    ///
    /// # Examples
    ///
    /// The following example demonstrates formatting headers during serialization.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_format2.rs
    /// #
    /// # use rust_xlsxwriter::{Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::Serialize;
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set a header format.
    ///     let header_format = Format::new()
    ///         .set_bold()
    ///         .set_border(FormatBorder::Thin)
    ///         .set_background_color("C6EFCE");
    ///
    ///     // Create a serializable struct.
    ///     #[derive(Serialize)]
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
    ///     let header_options = SerializeFieldOptions::new().set_header_format(header_format);
    ///
    ///     worksheet.serialize_headers_with_options(1, 1, &item1, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_format1.png">
    ///
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_header_format(mut self, format: impl Into<Format>) -> SerializeFieldOptions {
        self.header_format = Some(format.into());
        self
    }

    /// Hide all the headers.
    ///
    /// If you want to serialize data without outputting the headers above the
    /// data you can set the `hide_headers` parameters to any of the custom
    /// headers.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates turning off headers during serialization.
    /// The example in columns "D:E" have the headers turned off.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_format7.rs
    /// #
    /// # use rust_xlsxwriter::{SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a serializable struct.
    /// #     #[derive(Deserialize, Serialize)]
    /// #     struct Produce {
    /// #         fruit: &'static str,
    /// #         cost: f64,
    /// #     }
    /// #
    /// #     // Create some data instances.
    /// #     let items = [
    /// #         Produce {
    /// #             fruit: "Peach",
    /// #             cost: 1.05,
    /// #         },
    /// #         Produce {
    /// #             fruit: "Plum",
    /// #             cost: 0.15,
    /// #         },
    /// #         Produce {
    /// #             fruit: "Pear",
    /// #             cost: 0.75,
    /// #         },
    /// #     ];
    /// #
    ///     // Default serialization with headers (fruit and cost).
    ///     worksheet.deserialize_headers::<Produce>(0, 0)?;
    ///     worksheet.serialize(&items)?;
    ///
    ///     // Serialize the data but hide headers.
    ///     let header_options = SerializeFieldOptions::new().hide_headers(true);
    ///
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 3, &header_options)?;
    ///     worksheet.serialize(&items)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_format7.png">
    ///
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn hide_headers(mut self, enable: bool) -> SerializeFieldOptions {
        self.has_headers = !enable;
        self
    }

    /// Set custom properties for serialized fields.
    ///
    /// This method allows customization of of the serialization output of
    /// individual fields. It allows you to rename field headers, set
    /// formatting for serialized values, set the column width and other properties.
    ///
    /// See [`CustomSerializeField`] for more details.
    ///
    /// # Parameters
    ///
    /// - `custom_headers`: An array of [`CustomSerializeField`] values.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet with custom headers and cell
    /// formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_custom.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
    /// # };
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some formats to use with the serialization data.
    /// #     let header_format = Format::new()
    /// #         .set_bold()
    /// #         .set_border(FormatBorder::Thin)
    /// #         .set_background_color("C6EFCE");
    /// #
    /// #     let currency_format = Format::new().set_num_format("$0.00");
    /// #
    /// #     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
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
    ///     // Set up the custom headers.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("fruit")
    ///             .rename("Item"),
    ///         CustomSerializeField::new("cost")
    ///             .rename("Price")
    ///             .set_value_format(currency_format),
    ///     ];
    ///
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_header_format(header_format)
    ///         .set_custom_headers(&custom_headers);
    ///
    ///     // Set the serialization location and custom headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_custom.png">
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_custom_headers(
        mut self,
        custom_headers: &[CustomSerializeField],
    ) -> SerializeFieldOptions {
        self.custom_headers = custom_headers.to_vec();
        self
    }

    /// Add a default table structure to the serialized fields.
    ///
    /// This method can be used to add a default worksheet [`Table`] structure
    /// to a serialized area.
    ///
    /// See [`Table`] for more details on worksheet tables.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet and a default worksheet table
    /// structure.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_table1.rs
    /// #
    /// # use rust_xlsxwriter::{SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
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
    ///     // Set the header options.
    ///     let header_options = SerializeFieldOptions::new().set_table_default();
    ///
    ///     // Set the serialization location and custom headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_table1.png">
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_table_default(mut self) -> SerializeFieldOptions {
        self.table = Some(Table::new());
        self
    }

    /// Add a table structure to the serialized fields with a user defined style.
    ///
    /// This method can be used to add a  worksheet [`Table`] structure
    /// to a serialized area with a user defined [`TableStyle`].
    ///
    /// See [`Table`] for more details on worksheet tables.
    ///
    /// # Parameters
    ///
    /// - `style`: A [`TableStyle`] enum value.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde derived
    /// data structure to a worksheet with a worksheet table and a user defined
    /// style.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_table2.rs
    /// #
    /// # use rust_xlsxwriter::{SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
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
    ///     // Set the header options.
    ///     let header_options =
    ///         SerializeFieldOptions::new().set_table_style(rust_xlsxwriter::TableStyle::Medium10);
    ///
    ///     // Set the serialization location and custom headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/worksheet_serialize_table2.png">
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_table_style(mut self, style: TableStyle) -> SerializeFieldOptions {
        self.table = Some(Table::new().set_style(style));
        self
    }

    /// Add a user defined table structure to the serialized fields.
    ///
    /// This method can be used to add a user defined worksheet [`Table`]
    /// structure to a serialized area.
    ///
    /// See [`Table`] for more details on worksheet tables.
    ///
    /// # Parameters
    ///
    /// - `table`: A [`Table`] or object that implements `Into<Table>`.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet with a user defined worksheet
    /// table.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_table3.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     SerializeFieldOptions, Table, TableColumn, TableFunction, Workbook, XlsxError,
    /// # };
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
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
    ///     // Set the caption and subtotal in the total row.
    ///     let columns = vec![
    ///         TableColumn::new().set_total_label("Total"),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///     ];
    ///
    ///     // Create a new table and configure the total row.
    ///     let table = Table::new().set_total_row(true).set_columns(&columns);
    ///
    ///     // Set the header options.
    ///     let header_options = SerializeFieldOptions::new().set_table(table);
    ///
    ///     // Set the serialization location and custom headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_table3.png">
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_table(mut self, table: impl Into<Table>) -> SerializeFieldOptions {
        self.table = Some(table.into());
        self
    }

    /// Set the option to use only the specified custom headers.
    ///
    /// The default behavior when using custom headers is that the custom
    /// properties are merged over the default properties. So if you have a
    /// struct with three fields "item", "cost" and "availability" and you
    /// specify a custom header just for "cost" will still get all three fields
    /// output in the serialization.
    ///
    /// However, you may wish to output only the custom selected fields for use
    /// cases where you wish to skip fields or reorder them. The
    /// `use_custom_headers_only` property allow you to so that. See the
    /// example below for the effect of the option.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates different methods of handling custom
    /// properties. The user can either merge them with the default properties
    /// or use the custom properties exclusively.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_format8.rs
    /// #
    /// # use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a serializable struct.
    /// #     #[derive(Deserialize, Serialize)]
    /// #     struct Produce {
    /// #         fruit: &'static str,
    /// #         cost: f64,
    /// #         in_stock: bool,
    /// #     }
    /// #
    /// #     // Create some data instances.
    /// #     let items = [
    /// #         Produce {
    /// #             fruit: "Peach",
    /// #             cost: 1.05,
    /// #             in_stock: true,
    /// #         },
    /// #         Produce {
    /// #             fruit: "Plum",
    /// #             cost: 0.15,
    /// #             in_stock: false,
    /// #         },
    /// #         Produce {
    /// #             fruit: "Pear",
    /// #             cost: 0.75,
    /// #             in_stock: true,
    /// #         },
    /// #     ];
    /// #
    ///     // Default handling of customized headers: the formatting is merged with the
    ///     // default values so "in_stock" is still shown.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("fruit").rename("Item"),
    ///         CustomSerializeField::new("cost").rename("Price"),
    ///         CustomSerializeField::new("in_stock").rename("Foo"),
    ///     ];
    ///     let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);
    ///
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///     worksheet.serialize(&items)?;
    ///
    ///     // Set the "use_custom_headers_only" option to shown only the specified
    ///     // custom headers.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("fruit").rename("Item"),
    ///         CustomSerializeField::new("cost").rename("Price"),
    ///     ];
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_custom_headers(&custom_headers)
    ///         .use_custom_headers_only(true);
    ///
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 4, &header_options)?;
    ///     worksheet.serialize(&items)?;
    ///
    ///     // This can also be used to set the order of the output.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("cost").rename("Price"),
    ///         CustomSerializeField::new("fruit").rename("Item"),
    ///     ];
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_custom_headers(&custom_headers)
    ///         .use_custom_headers_only(true);
    ///
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 7, &header_options)?;
    ///     worksheet.serialize(&items)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_format8.png">
    ///
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn use_custom_headers_only(mut self, enable: bool) -> SerializeFieldOptions {
        self.use_custom_headers_only = enable;
        self
    }

    /// Set the name of the struct to be serialized.
    ///
    /// Note, this is a semi public method. End users won't and shouldn't need
    /// to use this since the struct name will be determine by the worksheet
    /// method that sets the header location. The method is here to allow full
    /// `SerializeFieldOptions` options to be generated programmatically by the
    /// [`XlsxSerialize`] proc macros.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the struct being serialized.
    ///
    #[doc(hidden)]
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_struct_name(mut self, name: impl Into<String>) -> SerializeFieldOptions {
        self.struct_name = name.into();
        self
    }
}

// -----------------------------------------------------------------------
// CustomSerializeField.
// -----------------------------------------------------------------------

/// The `CustomSerializeField` struct represents a custom serializer
/// field/header.
///
/// `CustomSerializeField` can be used to set column headers to map serialized
/// fields to. It allows you to rename, format or skip headers and also to define
/// formatting for field values.
///
/// It is used in conjunction with the [`SerializeFieldOptions`] struct and
/// [`Worksheet::deserialize_headers_with_options()`] and
/// [`Worksheet::serialize_headers_with_options()`] methods.
///
/// See [Working with Serde](crate::serializer) for an introduction to
/// serialization with `rust_xlsxwriter`.
///
/// # Examples
///
/// The following example demonstrates serializing instances of a Serde derived
/// data structure to a worksheet with custom headers and cell formatting.
///
/// ```
/// # // This code is available in examples/doc_worksheet_serialize_headers_custom.rs
/// #
/// use rust_xlsxwriter::{
///     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
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
///     let header_format = Format::new()
///         .set_bold()
///         .set_border(FormatBorder::Thin)
///         .set_background_color("C6EFCE");
///
///     let currency_format = Format::new().set_num_format("$0.00");
///
///     // Create a serializable struct.
///     #[derive(Deserialize, Serialize)]
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
///     // Set up the custom headers.
///     let custom_headers = [
///         CustomSerializeField::new("fruit")
///             .rename("Item"),
///         CustomSerializeField::new("cost")
///             .rename("Price")
///             .set_value_format(currency_format),
///     ];
///
///     let header_options = SerializeFieldOptions::new()
///         .set_header_format(header_format)
///         .set_custom_headers(&custom_headers);
///
///     // Set the serialization location and custom headers.
///     worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;
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
/// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_custom.png">
///
///
#[derive(Clone)]
#[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
pub struct CustomSerializeField {
    pub(crate) field_name: String,
    pub(crate) header_name: String,
    pub(crate) header_format: Option<Format>,
    pub(crate) column_format: Option<Format>,
    pub(crate) value_format: Arc<Option<Format>>,
    pub(crate) skip: bool,
    pub(crate) col: ColNum,
    pub(crate) width: Option<f64>,
    pub(crate) pixel_width: Option<u16>,
}

impl CustomSerializeField {
    /// Create custom serialize field/header options.
    ///
    /// Create a `CustomSerializeField` to be used with
    /// [`SerializeFieldOptions::set_custom_headers()`].
    ///
    /// The `field_name` argument must correspond to a struct field being
    /// serialized.
    ///
    /// # Parameters
    ///
    /// - `field_name`: The name of the serialized field to map to the header.
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn new(field_name: impl Into<String>) -> CustomSerializeField {
        let field_name = field_name.into();
        let header_name = field_name.clone();

        CustomSerializeField {
            field_name,
            header_name,
            header_format: None,
            column_format: None,
            value_format: Arc::new(None),
            skip: false,
            col: 0,
            width: None,
            pixel_width: None,
        }
    }

    /// Rename the field name displayed a custom serialize header.
    ///
    /// The field names of structs are serialized as column headers at the top
    /// of serialized data. The default field names may not be the header names
    /// that you want displayed in Excel in which case you can use one of the
    /// two main methods to rename the fields/headers:
    ///
    /// 1. Rename the field during serialization using the Serde:
    ///    - [field attribute]: `#[serde(rename = "name")` or
    ///    - [container attribute]: `#[serde(rename_all = "...")]`.
    /// 2. Rename the header (not field) when setting up custom serialization
    ///    headers via [`Worksheet::deserialize_headers_with_options()`] or
    ///    [`Worksheet::serialize_headers_with_options()`] and
    ///    [`CustomSerializeField::rename()`].
    ///
    /// [field attribute]: https://serde.rs/field-attrs.html
    /// [container attribute]: https://serde.rs/container-attrs.html
    ///
    /// See [Renaming fields when
    /// serializing](crate::serializer#renaming-fields-when-serializing) for
    /// more details.
    ///
    /// # Parameters
    ///
    /// - `name`: A string like name to use as the header.
    ///
    /// # Examples
    ///
    /// The following example demonstrates renaming fields during serialization
    /// by specifying custom headers and renaming them there. You must still
    /// specify the actual field name to serialize in the `new()` constructor.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_rename2.rs
    /// #
    /// # use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
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
    ///     // Set up the custom headers.
    ///     let custom_headers = [
    ///         CustomSerializeField::new("fruit").rename("Item"),
    ///         CustomSerializeField::new("cost").rename("Price"),
    ///     ];
    ///     let header_options = SerializeFieldOptions::new().set_custom_headers(&custom_headers);
    ///
    ///     // Set the serialization location and custom headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_rename1.png">
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn rename(mut self, name: impl Into<String>) -> CustomSerializeField {
        self.header_name = name.into();
        self
    }

    /// Set the header format for a custom serialize header.
    ///
    /// In general the [`SerializeFieldOptions::set_header_format()`] method
    /// should be used to set the header format for all custom headers but if
    /// you require individual headers to have different header formats you can
    /// use this method.
    ///
    /// See [`Format`] for more information on formatting.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`Format`] property for the custom header.
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_header_format(mut self, format: impl Into<Format>) -> CustomSerializeField {
        self.header_format = Some(format.into());
        self
    }

    /// Set the format for the column corresponding to a serialize header/field.
    ///
    /// This method can be used to set a number format, or other properties, for
    /// data that is serialized below the header. This method sets the format
    /// for the entire column whereas the
    /// [`CustomSerializeField::set_value_format()`] method below only sets it
    /// for serialized data within the column.
    ///
    /// This a a wrapper around the [`Worksheet::set_column_format()`] method
    /// with the advantage that it doesn't require you to keep track of the
    /// actual column number to use it.
    ///
    /// See [`Format`] and [`Format::set_num_format()`] for more information on
    /// formatting.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`Format`] property for the conditional format.
    ///
    /// # Examples
    ///
    /// The following example demonstrates serializing instances of a Serde
    /// derived data structure to a worksheet with header and column formatting.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_format5.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     CustomSerializeField, Format, FormatBorder, SerializeFieldOptions, Workbook, XlsxError,
    /// # };
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add some formats to use with the serialization data.
    ///     let header_format = Format::new()
    ///         .set_bold()
    ///         .set_border(FormatBorder::Thin)
    ///         .set_background_color("C6EFCE");
    ///
    ///     let currency_format = Format::new().set_num_format("$0.00");
    /// #
    /// #     // Create a serializable struct.
    /// #     #[derive(Deserialize, Serialize)]
    /// #     struct Produce {
    /// #         #[serde(rename = "Item")]
    /// #         fruit: &'static str,
    /// #
    /// #         #[serde(rename = "Price")]
    /// #         cost: f64,
    /// #     }
    /// #
    /// #     // Create some data instances.
    /// #     let item1 = Produce {
    /// #         fruit: "Peach",
    /// #         cost: 1.05,
    /// #     };
    /// #
    /// #     let item2 = Produce {
    /// #         fruit: "Plum",
    /// #         cost: 0.15,
    /// #     };
    /// #
    /// #     let item3 = Produce {
    /// #         fruit: "Pear",
    /// #         cost: 0.75,
    /// #     };
    /// #
    ///     // Set up the custom headers.
    ///     let custom_headers =
    ///         [CustomSerializeField::new("Price").set_column_format(currency_format)];
    ///
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_header_format(header_format)
    ///         .set_custom_headers(&custom_headers);
    ///
    ///     // Set the serialization location and headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(1, 1, &header_options)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_custom.png">
    ///
    pub fn set_column_format(mut self, format: impl Into<Format>) -> CustomSerializeField {
        self.column_format = Some(format.into());
        self
    }

    /// Set the cell format for values corresponding to a serialize
    /// header/field.
    ///
    /// This method can be used to set a number format, or other properties, for
    /// data that is serialized below the header. This method sets the format
    /// for the serialized data within the column whereas the
    /// [`CustomSerializeField::set_column_format()`] method above sets it for
    /// the the entire column.
    ///
    /// See [`Format`] and [`Format::set_num_format()`] for more information on
    /// formatting.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`Format`] property for cells corresponding to the
    ///   field/header.
    ///
    /// # Examples
    ///
    /// The following example demonstrates formatting cells during
    /// serialization. Note the currency format for the `cost` cells.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_format3.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     CustomSerializeField, Format, SerializeFieldOptions, Workbook, XlsxError,
    /// # };
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set a cell value format.
    ///     let value_format = Format::new().set_num_format("$0.00");
    /// #
    /// #     // Create a serializable struct.
    /// #     #[derive(Deserialize, Serialize)]
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
    /// #
    /// #     let item2 = Produce {
    /// #         fruit: "Plum",
    /// #         cost: 0.15,
    /// #     };
    /// #
    /// #     let item3 = Produce {
    /// #         fruit: "Pear",
    /// #         cost: 0.75,
    /// #     };
    /// #
    ///     // Set the custom headers.
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_custom_headers(
    ///             &[CustomSerializeField::new("cost").set_value_format(value_format)]
    ///         );
    ///
    ///     // Set the serialization location and headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_format3.png">
    ///
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn set_value_format(mut self, format: impl Into<Format>) -> CustomSerializeField {
        self.value_format = Arc::new(Some(format.into()));
        self
    }

    /// Skip a field when serializing.
    ///
    /// When serializing a struct you may not want all of the fields to be
    /// serialized. For example the struct may contain internal fields that
    /// aren't of interest to the end user. There are several ways to skip
    /// fields:
    ///
    /// 1. Using the Serde [field attributes] `#[serde(skip)]`. This is the
    ///    simplest and best method.
    /// 2. Explicitly omitting the field when setting up custom serialization
    ///    headers This method is useful when you can't add any additional
    ///    attributes on the struct.
    /// 3. Marking the field as skippable via custom headers and this `skip()`
    ///    method. This is only required in a few edge cases where the previous
    ///    methods won't work.
    ///
    /// [field attributes]: https://serde.rs/field-attrs.html
    ///
    /// See [Skipping fields when
    /// serializing](crate::serializer#skipping-fields-when-serializing) for
    /// more details.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates skipping fields during serialization
    /// by explicitly skipping them via custom headers.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_serialize_headers_skip3.rs
    /// #
    /// # use rust_xlsxwriter::{CustomSerializeField, SerializeFieldOptions, Workbook, XlsxError};
    /// # use serde::{Deserialize, Serialize};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a serializable struct.
    ///     #[derive(Deserialize, Serialize)]
    ///     struct Produce {
    ///         fruit: &'static str,
    ///         cost: f64,
    ///         in_stock: bool,
    ///     }
    ///
    ///     // Create some data instances.
    ///     let item1 = Produce {
    ///         fruit: "Peach",
    ///         cost: 1.05,
    ///         in_stock: true,
    ///     };
    ///
    ///     let item2 = Produce {
    ///         fruit: "Plum",
    ///         cost: 0.15,
    ///         in_stock: true,
    ///     };
    ///
    ///     let item3 = Produce {
    ///         fruit: "Pear",
    ///         cost: 0.75,
    ///         in_stock: false,
    ///     };
    ///
    ///     // We only need to set a custom header for the field we want to skip.
    ///     let header_options = SerializeFieldOptions::new()
    ///         .set_custom_headers(&[CustomSerializeField::new("in_stock").skip(true)]);
    ///
    ///     // Set the serialization location and custom headers.
    ///     worksheet.deserialize_headers_with_options::<Produce>(0, 0, &header_options)?;
    ///
    ///     // Serialize the data.
    ///     worksheet.serialize(&item1)?;
    ///     worksheet.serialize(&item2)?;
    ///     worksheet.serialize(&item3)?;
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
    /// src="https://rustxlsxwriter.github.io/images/worksheet_serialize_headers_skip1.png">
    ///
    #[cfg_attr(docsrs, doc(cfg(feature = "serde")))]
    pub fn skip(mut self, enable: bool) -> CustomSerializeField {
        self.skip = enable;
        self
    }

    /// Set the width for the column corresponding to a serialize header/field.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// worksheet column in character units.
    ///
    /// This a a wrapper around the [`Worksheet::set_column_width()`] method
    /// with the advantage that it doesn't require you to keep track of the
    /// actual column number to use it.
    ///
    /// # Parameters
    ///
    /// - `width`: The row width in character units.
    ///
    pub fn set_column_width(mut self, width: impl Into<f64>) -> CustomSerializeField {
        self.width = Some(width.into());
        self
    }

    /// Set the width for the column corresponding to a serialize header/field.
    ///
    /// The `set_column_width_pixels()` method is used to change the default
    /// width of a worksheet column in pixel units.
    ///
    /// This a a wrapper around the [`Worksheet::set_column_width_pixels()`]
    /// method with the advantage that it doesn't require you to keep track of
    /// the actual column number to use it.
    ///
    /// # Parameters
    ///
    /// - `width`: The row width in character units.
    ///
    pub fn set_column_width_pixels(mut self, width: u16) -> CustomSerializeField {
        self.pixel_width = Some(width);
        self
    }
}

// -----------------------------------------------------------------------
// Worksheet Serializer. This is the implementation of the Serializer trait to
// serialized a serde derived struct to an Excel worksheet.
// -----------------------------------------------------------------------
#[allow(unused_variables)]
impl<'a> ser::Serializer for &'a mut Worksheet {
    #[doc(hidden)]
    type Ok = ();
    #[doc(hidden)]
    type Error = XlsxError;
    #[doc(hidden)]
    type SerializeSeq = Self;
    #[doc(hidden)]
    type SerializeTuple = Self;
    #[doc(hidden)]
    type SerializeTupleStruct = Self;
    #[doc(hidden)]
    type SerializeTupleVariant = Self;
    #[doc(hidden)]
    type SerializeMap = Self;
    #[doc(hidden)]
    type SerializeStruct = Self;
    #[doc(hidden)]
    type SerializeStructVariant = Self;

    // Serialize all the default number types that fit into Excel's f64 type.
    #[doc(hidden)]
    fn serialize_bool(self, data: bool) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_i8(self, data: i8) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_u8(self, data: u8) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_i16(self, data: i16) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_u16(self, data: u16) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_i32(self, data: i32) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_u32(self, data: u32) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_i64(self, data: i64) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_u64(self, data: u64) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_f32(self, data: f32) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    #[doc(hidden)]
    fn serialize_f64(self, data: f64) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    // Serialize strings types.
    #[doc(hidden)]
    fn serialize_str(self, data: &str) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    // Excel doesn't have a character type. Serialize a char as a
    // single-character string.
    #[doc(hidden)]
    fn serialize_char(self, data: char) -> Result<(), XlsxError> {
        self.serialize_str(&data.to_string())
    }

    // Excel doesn't have a type equivalent to a byte array.
    #[doc(hidden)]
    fn serialize_bytes(self, data: &[u8]) -> Result<(), XlsxError> {
        Ok(())
    }

    // Serialize Some(T) values.
    #[doc(hidden)]
    fn serialize_some<T>(self, data: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        data.serialize(self)
    }

    // Empty/None/Null values in Excel are ignored unless the cell has
    // formatting in which case they are handled as a "blank" cell. For all of
    // these cases we write an empty string and the worksheet writer methods
    // will handle it correctly based on context.

    #[doc(hidden)]
    fn serialize_none(self) -> Result<(), XlsxError> {
        self.serialize_str("")
    }

    #[doc(hidden)]
    fn serialize_unit(self) -> Result<(), XlsxError> {
        self.serialize_none()
    }

    #[doc(hidden)]
    fn serialize_unit_struct(self, _name: &'static str) -> Result<(), XlsxError> {
        self.serialize_none()
    }

    // Excel doesn't have an equivalent for the structure so we ignore it.
    #[doc(hidden)]
    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
    ) -> Result<(), XlsxError> {
        Ok(())
    }

    // Try to handle this as a single value.
    #[doc(hidden)]
    fn serialize_newtype_struct<T>(self, _name: &'static str, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(self)
    }

    // Excel doesn't have an equivalent for the structure so we ignore it.
    #[doc(hidden)]
    fn serialize_newtype_variant<T>(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut *self)?;
        Ok(())
    }

    // Compound types.
    //
    // The only compound types we map into the Excel data model are structs.

    // Structs are the main primary data type used to map data structures into
    // Excel.
    fn serialize_struct(
        self,
        name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeStruct, XlsxError> {
        // Store the struct type name to check against user defined structs.
        self.serializer_state.set_current_struct(name);

        self.serialize_map(Some(len))
    }

    #[doc(hidden)]
    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq, XlsxError> {
        Ok(self)
    }

    // Not used.
    #[doc(hidden)]
    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, XlsxError> {
        self.serialize_seq(Some(len))
    }

    // Not used.
    #[doc(hidden)]
    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, XlsxError> {
        self.serialize_seq(Some(len))
    }

    // Not used.
    #[doc(hidden)]
    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleVariant, XlsxError> {
        variant.serialize(&mut *self)?;
        Ok(self)
    }

    // The field/values of structs are treated as a map.
    #[doc(hidden)]
    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap, XlsxError> {
        Ok(self)
    }

    // Not used.
    #[doc(hidden)]
    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStructVariant, XlsxError> {
        variant.serialize(&mut *self)?;
        Ok(self)
    }
}

// The following impls deal with the serialization of compound types.
// Currently we only support/use SerializeStruct and SerializeSeq.

// Structs are the main sequence type used by `rust_xlsxwriter`.
#[doc(hidden)]
impl<'a> ser::SerializeStruct for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, key: &'static str, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        // Store field name to allow us to map to the correct header/column.
        self.serializer_state.current_field = key.to_string();

        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// We also serialize sequences to map vectors/arrays to Excel.
#[doc(hidden)]
impl<'a> ser::SerializeSeq for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    // Serialize a single element of the sequence.
    fn serialize_element<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// Serialize tuple sequences.
#[doc(hidden)]
impl<'a> ser::SerializeTuple for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_element<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// Serialize tuple struct sequences.
#[doc(hidden)]
impl<'a> ser::SerializeTupleStruct for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// Serialize tuple variant sequences.
#[doc(hidden)]
impl<'a> ser::SerializeTupleVariant for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// Serialize tuple map sequences.
#[doc(hidden)]
impl<'a> ser::SerializeMap for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_key<T>(&mut self, key: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        key.serialize(&mut **self)
    }

    fn serialize_value<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// Serialize struct variant sequences.
#[doc(hidden)]
impl<'a> ser::SerializeStructVariant for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, key: &'static str, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        key.serialize(&mut **self)?;
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// -----------------------------------------------------------------------
// SerializerHeader. A struct used to store header/field name during
// serialization of the headers.
// -----------------------------------------------------------------------
pub(crate) struct SerializerHeader {
    pub(crate) struct_name: String,
    pub(crate) field_names: Vec<String>,
}

// -----------------------------------------------------------------------
// Header Serializer. This is the a simplified implementation of the Serializer
// trait to capture the headers/field names only.
// -----------------------------------------------------------------------
#[allow(unused_variables)]
impl<'a> ser::Serializer for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;
    type SerializeSeq = Self;
    type SerializeTuple = Self;
    type SerializeTupleStruct = Self;
    type SerializeTupleVariant = Self;
    type SerializeMap = Self;
    type SerializeStruct = Self;
    type SerializeStructVariant = Self;

    // Serialize strings types to capture the field names but ignore all other
    // types.
    fn serialize_str(self, data: &str) -> Result<(), XlsxError> {
        self.field_names.push(data.to_string());
        Ok(())
    }

    // Store the struct type/name to allow us to disambiguate structs.
    fn serialize_struct(
        self,
        name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeStruct, XlsxError> {
        self.struct_name = name.to_string();
        self.serialize_map(Some(len))
    }

    // Ignore all other primitive types.
    fn serialize_bool(self, data: bool) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_i8(self, data: i8) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_u8(self, data: u8) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_i16(self, data: i16) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_u16(self, data: u16) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_i32(self, data: i32) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_u32(self, data: u32) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_i64(self, data: i64) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_u64(self, data: u64) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_f32(self, data: f32) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_f64(self, data: f64) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_char(self, data: char) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_bytes(self, data: &[u8]) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_some<T>(self, data: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn serialize_none(self) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_unit(self) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
    ) -> Result<(), XlsxError> {
        Ok(())
    }

    fn serialize_newtype_struct<T>(self, _name: &'static str, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn serialize_newtype_variant<T>(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq, XlsxError> {
        Ok(self)
    }

    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, XlsxError> {
        self.serialize_seq(Some(len))
    }

    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, XlsxError> {
        self.serialize_seq(Some(len))
    }

    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleVariant, XlsxError> {
        variant.serialize(&mut *self)?;
        Ok(self)
    }

    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap, XlsxError> {
        Ok(self)
    }

    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStructVariant, XlsxError> {
        variant.serialize(&mut *self)?;
        Ok(self)
    }
}

// We are only interested in Struct fields. Other compound types are ignored.
impl<'a> ser::SerializeStruct for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, key: &'static str, _value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        // Serialize the key/field name but ignore the values.
        key.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

impl<'a> ser::SerializeSeq for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_element<T>(&mut self, _value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

impl<'a> ser::SerializeTuple for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_element<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

impl<'a> ser::SerializeTupleStruct for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, _value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

impl<'a> ser::SerializeTupleVariant for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, _value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

impl<'a> ser::SerializeMap for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_key<T>(&mut self, _key: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn serialize_value<T>(&mut self, _value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

impl<'a> ser::SerializeStructVariant for &'a mut SerializerHeader {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, _key: &'static str, _value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// -----------------------------------------------------------------------
// Header Deserializer. This is the a simplified implementation of the
// Deserializer trait to capture the headers/field names only.
// -----------------------------------------------------------------------
pub(crate) struct DeSerializerHeader<'a> {
    pub(crate) struct_name: &'a mut &'static str,
    pub(crate) field_names: &'a mut &'static [&'static str],
}

impl<'de, 'a> Deserializer<'de> for DeSerializerHeader<'a> {
    type Error = XlsxError;

    fn deserialize_struct<V>(
        self,
        name: &'static str,
        fields: &'static [&'static str],
        _visitor: V,
    ) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        *self.struct_name = name;
        *self.field_names = fields;
        Err(XlsxError::SerdeError("Deserialization error".to_string()))
    }

    fn deserialize_any<V>(self, _visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        Err(XlsxError::SerdeError("Deserialization error".to_string()))
    }

    serde::forward_to_deserialize_any! {
        bool i8 i16 i32 i64 u8 u16 u32 u64 f32 f64 char str string bytes
        byte_buf option unit unit_struct newtype_struct seq tuple
        tuple_struct map enum identifier ignored_any
    }
}

pub(crate) fn deserialize_headers<'de, T>() -> SerializerHeader
where
    T: Deserialize<'de>,
{
    let mut struct_name = "";
    let mut field_names: &[&str] = &[""];

    // Ignore the deserialization return since we have set up all the
    // Deserializer methods (above) to return quickly/with an error.
    let _ = T::deserialize(DeSerializerHeader {
        struct_name: &mut struct_name,
        field_names: &mut field_names,
    });

    let struct_name = struct_name.to_string();
    let field_names = field_names.iter().map(|&s| s.to_string()).collect();

    SerializerHeader {
        struct_name,
        field_names,
    }
}

// -----------------------------------------------------------------------
// XlsxSerializer trait. Trait to map `#[xlsx()]` attributes to
// `SerializeFieldOptions` options.
// -----------------------------------------------------------------------

/// Trait to map `#[xlsx()]` attributes to `SerializeFieldOptions` options.
///
/// The trait is mainly used by the `rust_xlsxwriter_derive` proc macros to
/// convert from `#[xlsx()]` attributes to [`SerializeFieldOptions`] and
/// [`CustomSerializeField`] options to configure the output of Serde
/// serialization.
///
/// See [Working with Serde](crate::serializer#working-with-serde).
///
pub trait XlsxSerialize {
    /// Map `#[xlsx()]` attributes to [`SerializeFieldOptions`].
    fn to_serialize_field_options() -> SerializeFieldOptions;
}
