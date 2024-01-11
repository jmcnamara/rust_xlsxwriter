// Provides the 'XlsxSerialize' derive macro for the `rust_xlsxwriter` crate.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! TODO
//!

use proc_macro::TokenStream;
use quote::{quote, ToTokens};
use syn::{
    parse_macro_input, Attribute, Data, DeriveInput, Expr, Fields, LitFloat, LitInt, LitStr, Token,
};

#[proc_macro_derive(XlsxSerialize, attributes(xlsx, serde))]
#[allow(clippy::too_many_lines)]
/// TODO
///
/// Add docs on attributes.
///
///
pub fn excel_serialize_derive(input: TokenStream) -> TokenStream {
    let ast = parse_macro_input!(input as DeriveInput);
    let (impl_generics, type_generics, where_clause) = ast.generics.split_for_impl();
    let mut struct_name = ast.ident.to_string();
    let struct_type = ast.ident;

    let mut field_case = "original".to_string();
    let mut custom_fields = Vec::new();
    let mut field_options = quote!();
    let mut has_format_object = false;
    let mut format_use_statements = quote!();

    // Parse and handle container attributes.
    for attribute_tokens in &ast.attrs {
        for attribute in parse_header_attribute(attribute_tokens) {
            match attribute {
                // Handle container #[xlsx(field_options = "")] attribute.
                HeaderAttributeTypes::HideHeaders => {
                    field_options = quote! {
                        .hide_headers(true)
                    }
                }

                // Handle container #[xlsx(header_format = "")] attribute.
                HeaderAttributeTypes::HeaderFormat(format) => {
                    field_options = quote! {
                        #field_options
                        .set_header_format(#format)
                    };
                    has_format_object = true;
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
                                    has_format_object = true;
                                }

                                // Handle the #[xlsx(value_format = Format)] field attribute.
                                FieldAttributeTypes::ValueFormat(format) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_value_format(#format)
                                    };
                                    has_format_object = true;
                                }

                                // Handle the #[xlsx(column_format = Format)] field attribute.
                                FieldAttributeTypes::ColumnFormat(format) => {
                                    custom_field_methods = quote! {
                                        #custom_field_methods
                                        .set_column_format(#format)
                                    };
                                    has_format_object = true;
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

    // If the code includes Format::new() then provide some "use" statements.
    if has_format_object {
        format_use_statements = quote!(
            #[allow(unused_imports)]
            use ::rust_xlsxwriter::{
                Color, Format, FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern,
                FormatScript, FormatUnderline,
            };
        );
    }

    // Generate the impl for the derived struct. This creates a `SerializeFieldOptions`
    // struct and populates it with `CustomSerializeField` instances.
    let code = quote! {
        #[doc(hidden)]
        const _: () = {
            #format_use_statements
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
            // Handle the #[xlsx(header_format = Format)] container attribute.
            else if meta.path.is_ident("header_format") {
                let value = meta.value()?;
                let token = value.parse()?;
                attributes.push(HeaderAttributeTypes::HeaderFormat(token));
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
