// Provides the 'XlsxSerialize' derive macro for the `rust_xlsxwriter` crate.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The `rust_xlsxwriter_derive` crate provides the `XlsxSerialize` derived
//! trait which is used in conjunction with `rust_xlsxwriter` serialization.
//!
//! # Introduction
//!
//! [`XlsxSerialize`] can be used to set container and field attributes for
//! structs to define Excel formatting and other options when serializing them
//! to Excel using `rust_xlsxwriter` and `Serde`.
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
//!         #[xlsx(rename = "Price")]
//!         #[xlsx(value_format = Format::new().set_num_format("$0.00"))]
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
//! The output file is shown below. Note the change or column width in Column A,
//! the renamed headers and the currency format in Column B numbers.
//!
//! <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_intro.png">
//!
//! For more information see the documentation for [`XlsxSerialize`] or [Working
//! with Serde] in the `rust_xlsxwriter` docs.
//!
//! [Working with Serde]:
//!     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html
//!
use proc_macro::TokenStream;
use quote::{quote, ToTokens};
use syn::{
    parse_macro_input, Attribute, Data, DeriveInput, Expr, Fields, LitFloat, LitInt, LitStr, Token,
};

/// The `XlsxSerialize` derived trait is used in conjunction with
/// `rust_xlsxwriter` serialization.
///
/// # Introduction
///
/// `XlsxSerialize` can be used to set container and field attributes for
/// structs to define Excel formatting and other options when serializing them
/// to Excel using the `rust_xlsxwriter` crate.
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
/// The output file is shown below. Note the change or column width in Column A,
/// the renamed headers and the currency format in Column B numbers.
///
/// <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_intro.png">
///
/// For more information on serialization to Excel see [Working with Serde] in
/// the `rust_xlsxwriter` docs.
///
///
///
///
/// # The `xlsx` attributes
///
/// The `XlsxSerializer` derived trait adds a wrapper around a Serde
/// serializable struct to add `rust_xlsxwriter` specific formatting options. It
/// achieves this via the `rust_xlsxwriter` [`SerializeFieldOptions`] and
/// [`CustomSerializeField`] serialization configuration structs.
///
/// The attributes are divided into "Container" attributes that apply to the
/// entire struct and "Field" attributes which apply to individuals fields. In
/// an Excel context the field attributes apply to the serialization headers and
/// the data below them.
///
/// In order to demonstrate the each attribute and its effect in the next
/// sections we will use variations of the following example with the relevant
/// attribute applied.
///
/// ```
/// # // This code is available in examples/doc_xlsxserialize_base.rs
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
///     // Set the serialization location and headers.
///     worksheet.set_serialize_headers::<Produce>(0, 0)?;
///
///     // Serialize the data.
///     worksheet.serialize(&items)?;
///
///     // Save the file to disk.
///     workbook.save("serialize.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// ## Container `xlsx` attributes
///
/// The following are the "Container" attributes supported by `XlsxSerializer`:
///
/// - `#[xlsx(header_format = Format)`
///
///   The `header_format` container attribute sets the [`Format`] property for
///   headers. See the [Working with attribute
///   Formats](#working-with-attribute-formats) section below for information on
///   handling formats.
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         #[xlsx(header_format = Format::new()
///                .set_bold()
///                .set_border(FormatBorder::Thin)
///                .set_background_color("C6EFCE"))]
///         struct Produce {
///             fruit: &'static str,
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_header_format.png">
///
///
///
/// - `#[xlsx(hide_headers)`
///
///   The `hide_headers` container attribute hides the serialization headers:
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         #[xlsx(hide_headers)]
///         struct Produce {
///             fruit: &'static str,
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_hide_headers.png">
///
///
/// - `#[xlsx(table_default)`
///
///   The `table_default` container attribute adds a worksheet [`Table`]
///   structure with default formatting to the serialized area:
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         #[xlsx(table_default)]
///         struct Produce {
///             fruit: &'static str,
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_table_default.png">
///
///
/// - `#[xlsx(table_style)`
///
///   The `table_style` container attribute adds a worksheet [`Table`]
///   structure with a user specified [`TableStyle`] to the serialized area:
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         #[xlsx(table_style = TableStyle::Medium10)]
///         struct Produce {
///             fruit: &'static str,
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_table_style.png">
///
///
/// - `#[xlsx(table = Table)`
///
///   The `table` container attribute adds a user defined worksheet [`Table`]
///   structure to the serialized area:
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         #[xlsx(table = Table::new())]
///         struct Produce {
///             fruit: &'static str,
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_table_default.png">
///
///   See the [Working with attribute
///   Formats](#working-with-attribute-formats) section below for information on
///   how to wrap complex objects like [`Format`] or [`Table`] in a function so
///   it can be used as an attribute parameter.
///
///
/// ## Field `xlsx` attributes
///
/// The following are the "Field" attributes supported by `XlsxSerializer`:
///
/// - `#[xlsx(rename = "")`
///
///   The `rename` field attribute renames the Excel header. This is similar to
///   the `#[serde(rename = "")` field attribute except that it doesn't rename
///   the field for other types of serialization.
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         struct Produce {
///             #[xlsx(rename = "Item")]
///             fruit: &'static str,
///
///             #[xlsx(rename = "Price")]
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_rename.png">
///
///
/// - `#[xlsx(num_format = "")`
///
///   The `num_format` field attribute sets the property to change the number
///   formatting of the output. It is a syntactic shortcut for
///   [`Format::set_num_format()`], see below.
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         struct Produce {
///             fruit: &'static str,
///
///             #[xlsx(num_format = "$0.00")]
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_num_format.png">
///
///
/// - `#[xlsx(value_format = Format)`
///
///   The `value_format` field attribute sets the [`Format`] property for the
///   fields below the header.
///
///   See the [Working with attribute Formats](#working-with-attribute-formats)
///   section below for information on handling formats.
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         struct Produce {
///             #[xlsx(value_format = Format::new().set_font_color("#FF0000"))]
///             fruit: &'static str,
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_value_format.png">
///
///
/// - `#[xlsx(column_format = Format)`
///
///   The `column_format` field attribute is similar to the previous
///   `value_format` attribute except that it sets the format for the entire
///   column, including any data added manually below the serialized data.
///
///   <br>
///
/// - `#[xlsx(header_format = Format)`
///
///   The `header_format` field attribute sets the [`Format`] property for the
///   the header. It is similar to the container method of the same name except
///   it only sets the format or one header:
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         struct Produce {
///             fruit: &'static str,
///
///             #[xlsx(header_format = Format::new().set_bold().set_font_color("#FF0000"))]
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_field_header_format.png">
///
///
/// - `#[xlsx(column_width = float)`
///
///   The `column_width` field attribute sets the column width in character
///   units.
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         struct Produce {
///             fruit: &'static str,
///             #[xlsx(column_width = 20.0)]
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img
///   src="https://rustxlsxwriter.github.io/images/xlsxserialize_column_width.png">
///
///
///
/// - `#[xlsx(column_width_pixels = int)`
///
///   The `column_width_pixels` field attribute field attribute is similar to
///   the previous `column_width` attribute except that the width is specified
///   in integer pixel units.
///
///   <br>
///
///
///
///
/// - `#[xlsx(skip)`
///
///   The `skip` field attribute skips writing the field to the target Excel
///   file.  This is similar to the `#[serde(skip)` field attribute except that
///   it doesn't skip the field for other types of serialization.
///
///   ```
///   # use rust_xlsxwriter::XlsxSerialize;
///   # use serde::Serialize;
///   #
///   # fn main() {
///         #[derive(XlsxSerialize, Serialize)]
///         struct Produce {
///             fruit: &'static str,
///
///             #[xlsx(skip)]
///             cost: f64,
///         }
///   # }
///   ```
///
///   <img src="https://rustxlsxwriter.github.io/images/xlsxserialize_skip.png">
///
///
/// Note, if required you can group more than one attribute
///
/// ```
/// #[xlsx(rename = "Item", column_width = 20.0)]
/// ```
///
/// ## Working with attribute Formats
///
/// When working with `XlsxSerialize` attributes that deal with [`Format`]
/// objects the definition can become quite long:
///
/// ```
/// # use rust_xlsxwriter::XlsxSerialize;
/// # use serde::Serialize;
/// #
/// # fn main() {
///       #[derive(XlsxSerialize, Serialize)]
///       #[xlsx(header_format = Format::new()
///              .set_bold()
///              .set_border(FormatBorder::Thin)
///              .set_background_color("C6EFCE"))]
///       struct Produce {
///           fruit: &'static str,
///           cost: f64,
///       }
/// # }
/// ```
///
/// You might in this case be tempted to define the format in another part of
/// your code and use a variable to define the format:
///
/// ```compile_fail
/// # use rust_xlsxwriter::{Format, FormatBorder, XlsxSerialize};
/// # use serde::Serialize;
/// #
/// # fn main() {
///     let my_header_format = Format::new()
///         .set_bold()
///         .set_border(FormatBorder::Thin)
///         .set_background_color("C6EFCE");
///
///     #[derive(XlsxSerialize, Serialize)]
///     #[xlsx(header_format = &my_header_format)] // Won't work.
///     struct Produce {
///         fruit: &'static str,
///         cost: f64,
///     }
/// # }
/// ```
///
/// <br>
///
/// However, this won't work because Rust derive/proc macros are compiled
/// statically and `&header_format` is a dynamic variable.
///
/// A workaround for this it to define any formats you wish to use in functions:
///
/// ```
/// # // This code is available in examples/doc_xlsxserialize_header_format_reuse.rs
/// #
/// # use rust_xlsxwriter::{FormatBorder, Workbook, XlsxError, XlsxSerialize, Format};
/// # use serde::Serialize;
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
///     fn my_header_format() -> Format {
///         Format::new()
///             .set_bold()
///             .set_border(FormatBorder::Thin)
///             .set_background_color("C6EFCE")
///     }
///
///     #[derive(XlsxSerialize, Serialize)]
///     #[xlsx(header_format = my_header_format())]
///     struct Produce {
///         fruit: &'static str,
///         cost: f64,
///     }
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
/// #     // Set the serialization location and headers.
/// #     worksheet.set_serialize_headers::<Produce>(0, 0)?;
/// #
/// #     // Serialize the data.
/// #     worksheet.serialize(&items)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("serialize.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/xlsxserialize_header_format_reuse.png">
///
///
///
///
/// [Serde Attributes]: https://serde.rs/attributes.html
///
/// [`Format`]:
///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html
///
/// [`Table`]:
///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Table.html
///
/// [`TableStyle`]:
///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.TableStyle.html
///
/// [`Format::set_num_format()`]:
///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#method.set_num_format
///
/// [`SerializeFieldOptions`]:
///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/struct.SerializeFieldOptions.html
///
/// [`CustomSerializeField`]:
///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/struct.CustomSerializeField.html
///
///
///
///
///
#[proc_macro_derive(XlsxSerialize, attributes(xlsx, serde))]
#[allow(clippy::too_many_lines)]
pub fn excel_serialize_derive(input: TokenStream) -> TokenStream {
    let ast = parse_macro_input!(input as DeriveInput);
    let (impl_generics, type_generics, where_clause) = ast.generics.split_for_impl();
    let mut struct_name = ast.ident.to_string();
    let struct_type = ast.ident;

    let mut field_case = "original".to_string();
    let mut custom_fields = Vec::new();
    let mut field_options = quote!();
    let mut has_includes = false;
    let mut use_statements = quote!();

    // Parse and handle container attributes.
    for attribute_tokens in &ast.attrs {
        for attribute in parse_header_attribute(attribute_tokens) {
            match attribute {
                // Handle container #[xlsx(hide_headers)] attribute.
                HeaderAttributeTypes::HideHeaders => {
                    field_options = quote! {
                        #field_options
                        .hide_headers(true)
                    };
                }

                // Handle container #[xlsx(table_default)] attribute.
                HeaderAttributeTypes::TableDefault => {
                    field_options = quote! {
                        #field_options
                        .set_table_default()
                    };
                    has_includes = true;
                }

                // Handle container #[xlsx(header_format = "")] attribute.
                HeaderAttributeTypes::HeaderFormat(format) => {
                    field_options = quote! {
                        #field_options
                        .set_header_format(#format)
                    };
                    has_includes = true;
                }

                // Handle container #[xlsx(table_style = "")] attribute.
                HeaderAttributeTypes::TableStyle(style) => {
                    field_options = quote! {
                        #field_options
                        .set_table_style(#style)
                    };
                    has_includes = true;
                }

                // Handle container #[xlsx(table = "")] attribute.
                HeaderAttributeTypes::Table(table) => {
                    field_options = quote! {
                        #field_options
                        .set_table(#table)
                    };
                    has_includes = true;
                }

                // Handle container #[serde(rename = "")] attribute.
                HeaderAttributeTypes::SerdeRename(name) => {
                    struct_name = name.value();
                }

                // Handle container #[serde(rename_all = "")] attribute.
                HeaderAttributeTypes::SerdeRenameAll(name) => field_case = name.value(),

                // Raise any errors from parsing the attributes.
                HeaderAttributeTypes::Error(error_code) => {
                    return error_code;
                }
            }
        }
    }

    // Parse and handle field attributes.
    if let Data::Struct(data) = ast.data {
        if let Fields::Named(fields) = data.fields {
            'field: for field in &fields.named {
                if let Some(field_name) = field.ident.as_ref() {
                    // Get the field name to map to a custom header.
                    let mut field_name = field_name.to_string();

                    if field_name != "original" {
                        field_name = rename_field(&field_name, &field_case);
                    }

                    let mut custom_field_constructor = quote! {
                        ::rust_xlsxwriter::CustomSerializeField::new(#field_name)
                    };

                    let mut custom_field_methods = quote! {};

                    for attribute_tokens in &field.attrs {
                        for attribute in parse_field_attribute(attribute_tokens) {
                            match attribute {
                                // Handle the #[xlsx(rename = "")] field attribute. This is different
                                // from serde "rename" since it doesn't rename the struct field
                                // just the string in Excel.
                                FieldAttributeTypes::Rename(name) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .rename(#name)
                                    };
                                }

                                // Handle the #[xlsx(header_format = Format)] field attribute.
                                FieldAttributeTypes::HeaderFormat(format) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_header_format(#format)
                                    };
                                    has_includes = true;
                                }

                                // Handle the #[xlsx(value_format = Format)] field attribute.
                                FieldAttributeTypes::ValueFormat(format) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_value_format(#format)
                                    };
                                    has_includes = true;
                                }

                                // Handle the #[xlsx(column_format = Format)] field attribute.
                                FieldAttributeTypes::ColumnFormat(format) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_column_format(#format)
                                    };
                                    has_includes = true;
                                }

                                // Handle the #[xlsx(num_format = "")] field attribute.
                                FieldAttributeTypes::NumFormat(num_format) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_value_format(#num_format)
                                    };
                                }

                                // Handle the #[xlsx(column_width = float)] field attribute.
                                FieldAttributeTypes::ColumnWidth(width) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_column_width(#width)
                                    };
                                }

                                // Handle the #[xlsx(column_width_pixels = int)] field attribute.
                                FieldAttributeTypes::ColumnWidthPixels(width) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_column_width_pixels(#width)
                                    };
                                }

                                // Handle the #[xlsx(skip)] field attribute by setting the
                                // .skip() property of the custom header.
                                FieldAttributeTypes::Skip => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .skip(true)
                                    };
                                }

                                // Handle the #[serde(rename = "")] field attribute.
                                FieldAttributeTypes::SerdeRename(field_name) => {
                                    custom_field_constructor = quote! {
                                        ::rust_xlsxwriter::CustomSerializeField::new(#field_name)
                                    };
                                }

                                // Handle the #[serde(skip)] field attribute attribute by ignoring
                                // the field.
                                FieldAttributeTypes::SerdeSkip => {
                                    continue 'field;
                                }

                                // Raise any errors from parsing the attributes.
                                FieldAttributeTypes::Error(error_code) => {
                                    return error_code;
                                }
                            }
                        }
                    }

                    let custom_field = quote! {
                        #custom_field_constructor
                        #custom_field_methods
                    };

                    custom_fields.push(custom_field);
                }
            }
        }
    }

    // If the code includes Format::new() or Table::new() then provide some
    // "use" statements.
    if has_includes {
        use_statements = quote!(
            #[allow(unused_imports)]
            use ::rust_xlsxwriter::{
                Color, Format, FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern,
                FormatScript, FormatUnderline, Table, TableColumn, TableFunction, TableStyle,
            };
        );
    }

    // Generate the impl for the derived struct. This creates a `SerializeFieldOptions`
    // struct and populates it with `CustomSerializeField` instances.
    let code = quote! {
        #[doc(hidden)]
        const _: () = {
            #use_statements
            impl #impl_generics ::rust_xlsxwriter::XlsxSerialize for #struct_type #type_generics #where_clause {
                fn to_serialize_field_options() -> ::rust_xlsxwriter::SerializeFieldOptions {
                    let custom_headers = [
                        #( #custom_fields ),*
                    ];

                    ::rust_xlsxwriter::SerializeFieldOptions::new()
                        #field_options
                        .set_struct_name(#struct_name)
                        .set_custom_headers(&custom_headers)
                }
            }
        };
    };
    code.into()
}

// Parse the container attributes for `xlsx` and *some* `serde` options.
//
// Example:
//
// ```
// #[derive(XlsxSerialize, Serialize)]
// #[xlsx(hide_headers)]
// #[serde(rename = "MyStruct")]
// #[serde(rename_all = "PascalCase")]
// struct Produce {
//     fruit: &'static str,
//     cost: f64,
//     in_stock: bool,
// }
// ```
//
fn parse_header_attribute(attribute: &Attribute) -> Vec<HeaderAttributeTypes> {
    let mut attributes = vec![];

    if attribute.path().is_ident("xlsx") {
        let parse_result = attribute.parse_nested_meta(|meta| {
            // Handle the #[xlsx(hide_headers)] container attribute.
            if meta.path.is_ident("hide_headers") {
                attributes.push(HeaderAttributeTypes::HideHeaders);
                Ok(())
            }
            // Handle the #[xlsx(table_default)] container attribute.
            else if meta.path.is_ident("table_default") {
                attributes.push(HeaderAttributeTypes::TableDefault);
                Ok(())
            }
            // Handle the #[xlsx(header_format = Format)] container attribute.
            else if meta.path.is_ident("header_format") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(HeaderAttributeTypes::HeaderFormat(token));
                Ok(())
            }
            // Handle the #[xlsx(table_style = TableStyle::*)] container attribute.
            else if meta.path.is_ident("table_style") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(HeaderAttributeTypes::TableStyle(token));
                Ok(())
            }
            // Handle the #[xlsx(table = TableStyle::new())] container attribute.
            else if meta.path.is_ident("table") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(HeaderAttributeTypes::Table(token));
                Ok(())
            }
            // Handle any unrecognized attributes as an error.
            else {
                let path = meta.path.to_token_stream().to_string();
                let message = format!("unknown rust_xlsxwriter xlsx attribute: `{path}`");
                Err(meta.error(message))
            }
        });

        if let Err(err) = parse_result {
            let error = err.into_compile_error();
            attributes.push(HeaderAttributeTypes::Error(error.into()));
        }
    }

    // Limited handling of Serde attributes. We don't try to catch or handle any
    // errors since that will be done by the Serde proc macros.
    if attribute.path().is_ident("serde") {
        let _ = attribute.parse_nested_meta(|meta| {
            // We need to handle 2 `rename_all` cases here, one of which is nested:
            //     #[serde(rename_all = "...")]
            //     #[serde(rename_all(serialize = "..."))]
            if meta.path.is_ident("rename_all") {
                let not_nested = meta.input.peek(Token![=]);

                if not_nested {
                    let value = meta.value()?;
                    let token = value.parse()?;
                    attributes.push(HeaderAttributeTypes::SerdeRenameAll(token));
                } else {
                    let _ = meta.parse_nested_meta(|meta| {
                        if meta.path.is_ident("serialize") {
                            let value = meta.value()?;
                            let token = value.parse()?;
                            attributes.push(HeaderAttributeTypes::SerdeRenameAll(token));
                        }
                        Ok(())
                    });
                }

                Ok(())
            }
            // Handle the #[serde(rename = "")] container attribute.
            else if meta.path.is_ident("rename") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(HeaderAttributeTypes::SerdeRename(token));
                Ok(())
            }
            // Ignore everything else.
            else {
                Ok(())
            }
        });
    }

    attributes
}

// Header attribute return values.
enum HeaderAttributeTypes {
    Error(TokenStream),
    HideHeaders,
    HeaderFormat(Expr),
    TableDefault,
    TableStyle(Expr),
    Table(Expr),
    SerdeRename(LitStr),
    SerdeRenameAll(LitStr),
}

// Parse the field attributes for `xlsx` and *some* `serde` options.
//
// Example:
//
// ```
// #[derive(XlsxSerialize, Serialize)]
// struct Produce {
//     #[serde(rename = "Item")]
//     fruit: &'static str,
//
//     #[xlsx(rename = "Price")]
//     #[xlsx(num_format = "$0.00")]
//     #[xlsx(column_width = 10.0)]
//     cost: f64,
//
//     #[serde(skip)]
//     in_stock: bool,
// }
// ```
//
fn parse_field_attribute(attribute: &Attribute) -> Vec<FieldAttributeTypes> {
    let mut attributes = vec![];

    if attribute.path().is_ident("xlsx") {
        let parse_result = attribute.parse_nested_meta(|meta| {
            // Handle the #[xlsx(rename = "")] field attribute.
            if meta.path.is_ident("rename") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::Rename(token));
                Ok(())
            }
            // Handle the #[xlsx(num_format = "")] field attribute.
            else if meta.path.is_ident("num_format") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::NumFormat(token));
                Ok(())
            }
            // Handle the #[xlsx(header_format = Format)] field attribute.
            else if meta.path.is_ident("header_format") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::HeaderFormat(token));
                Ok(())
            }
            // Handle the #[xlsx(value_format = Format)] field attribute.
            else if meta.path.is_ident("value_format") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::ValueFormat(token));
                Ok(())
            }
            // Handle the #[xlsx(column_format = Format)] field attribute.
            else if meta.path.is_ident("column_format") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::ColumnFormat(token));
                Ok(())
            }
            // Handle the #[xlsx(column_width = float)] field attribute.
            else if meta.path.is_ident("column_width") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::ColumnWidth(token));
                Ok(())
            }
            // Handle the #[xlsx(column_width_pixels = int)] field attribute.
            else if meta.path.is_ident("column_width_pixels") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::ColumnWidthPixels(token));
                Ok(())
            }
            // Handle the #[xlsx(skip)] field attribute.
            else if meta.path.is_ident("skip") {
                attributes.push(FieldAttributeTypes::Skip);
                Ok(())
            }
            // Handle any unrecognized attributes as an error.
            else {
                let path = meta.path.to_token_stream().to_string();
                let message = format!("unknown rust_xlsxwriter xlsx attribute: `{path}`");
                Err(meta.error(message))
            }
        });

        if let Err(err) = parse_result {
            let error = err.into_compile_error();
            attributes.push(FieldAttributeTypes::Error(error.into()));
        }
    }

    // Limited handling of Serde attributes. We don't try to catch or handle any
    // errors since that will be done by the Serde proc macros.
    if attribute.path().is_ident("serde") {
        let _ = attribute.parse_nested_meta(|meta| {
            // Handle the serde `skip` field attributes:
            //    #[serde(skip)]
            //    #[serde(skip_serializing)]
            if meta.path.is_ident("skip") || meta.path.is_ident("skip_serializing") {
                attributes.push(FieldAttributeTypes::SerdeSkip);
                Ok(())
            }
            // Handle he #[serde(rename = "Price")] field attribute:
            else if meta.path.is_ident("rename") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(FieldAttributeTypes::SerdeRename(token));
                Ok(())
            }
            // Ignore everything else.
            else {
                Ok(())
            }
        });
    }

    attributes
}

// Field attribute return values.
enum FieldAttributeTypes {
    Skip,
    Error(TokenStream),
    Rename(LitStr),
    NumFormat(LitStr),
    HeaderFormat(Expr),
    ValueFormat(Expr),
    ColumnFormat(Expr),
    ColumnWidth(LitFloat),
    ColumnWidthPixels(LitInt),
    SerdeSkip,
    SerdeRename(LitStr),
}

// -----------------------------------------------------------------------
// Function to mimic Serde's RenameRule.apply_to_field().
// -----------------------------------------------------------------------
fn rename_field(field_name: &str, rename_type: &str) -> String {
    match rename_type {
        "lowercase" => field_name.to_ascii_lowercase(),
        "camelCase" => camel_case(field_name),
        "kebab-case" => field_name.replace('_', "-"),
        "PascalCase" => pascal_case(field_name),
        "SCREAMING-KEBAB-CASE" => field_name.replace('_', "-").to_ascii_uppercase(),
        "UPPERCASE" | "SCREAMING_SNAKE_CASE" => field_name.to_ascii_uppercase(),
        _ => field_name.to_string(),
    }
}

fn pascal_case(field_name: &str) -> String {
    field_name
        .split('_')
        .map(uppercase_first)
        .collect::<String>()
}

fn camel_case(field_name: &str) -> String {
    lowercase_first(&pascal_case(field_name))
}

fn uppercase_first(segment: &str) -> String {
    let mut segment = segment.to_string();
    segment.remove(0).to_uppercase().to_string() + &segment
}

fn lowercase_first(segment: &str) -> String {
    let mut segment = segment.to_string();
    segment.remove(0).to_lowercase().to_string() + &segment
}

// -----------------------------------------------------------------------
// Test input taken from a Serde test case.
// -----------------------------------------------------------------------
#[test]
fn rename_fields() {
    for (input, case_type, expected) in [
        // Test data 1.
        ("outcome", "original", "outcome"),
        ("outcome", "UPPERCASE", "OUTCOME"),
        ("outcome", "PascalCase", "Outcome"),
        ("outcome", "camelCase", "outcome"),
        ("outcome", "SCREAMING_SNAKE_CASE", "OUTCOME"),
        ("outcome", "kebab-case", "outcome"),
        ("outcome", "SCREAMING-KEBAB-CASE", "OUTCOME"),
        // Test data 2.
        ("very_tasty", "original", "very_tasty"),
        ("very_tasty", "UPPERCASE", "VERY_TASTY"),
        ("very_tasty", "PascalCase", "VeryTasty"),
        ("very_tasty", "camelCase", "veryTasty"),
        ("very_tasty", "SCREAMING_SNAKE_CASE", "VERY_TASTY"),
        ("very_tasty", "kebab-case", "very-tasty"),
        ("very_tasty", "SCREAMING-KEBAB-CASE", "VERY-TASTY"),
        // Test data 3.
        ("a", "original", "a"),
        ("a", "UPPERCASE", "A"),
        ("a", "PascalCase", "A"),
        ("a", "camelCase", "a"),
        ("a", "SCREAMING_SNAKE_CASE", "A"),
        ("a", "kebab-case", "a"),
        ("a", "SCREAMING-KEBAB-CASE", "A"),
        // Test data 4.
        ("z42", "original", "z42"),
        ("z42", "UPPERCASE", "Z42"),
        ("z42", "PascalCase", "Z42"),
        ("z42", "camelCase", "z42"),
        ("z42", "SCREAMING_SNAKE_CASE", "Z42"),
        ("z42", "kebab-case", "z42"),
        ("z42", "SCREAMING-KEBAB-CASE", "Z42"),
    ] {
        assert_eq!(
            expected,
            &rename_field(input, case_type),
            "for {}",
            case_type
        );
    }
}
