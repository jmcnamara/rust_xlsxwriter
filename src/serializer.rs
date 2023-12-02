// serializer - A serde serializer for use with `rust_xlsxwriter`.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

mod error {
    pub use serde::de::value::Error;
}

use std::collections::HashMap;

use crate::{ColNum, IntoExcelData, RowNum, Worksheet, XlsxError};
use error::Error;
use serde::{ser, Serialize};

const MAX_LOSSLESS_U64_TO_F64: u64 = 2 << 52;
const MAX_LOSSLESS_I64_TO_F64: i64 = 2 << 52;
const MIN_LOSSLESS_I64_TO_F64: i64 = -MAX_LOSSLESS_I64_TO_F64;

// -----------------------------------------------------------------------

pub(crate) struct SerializerState {
    headers: HashMap<String, SerializerHeader>,
    current_field: String,
    current_col: ColNum,
    current_row: RowNum,
}

impl SerializerState {
    // TODO
    pub(crate) fn new() -> SerializerState {
        SerializerState {
            headers: HashMap::new(),
            current_field: String::new(),
            current_col: 0,
            current_row: 0,
        }
    }

    // TODO
    pub(crate) fn set_row_col_for_field(&mut self, key: &'static str) -> Result<(), Error> {
        let Some(header) = self.headers.get_mut(key) else {
            return Err(serde::de::Error::custom(format!(
                "unknown field '{key}', add it via Worksheet::add_serialize_headers()"
            )));
        };

        header.row += 1;

        self.current_col = header.col;
        self.current_row = header.row;
        self.current_field = key.to_string();

        Ok(())
    }

    // TODO
    pub(crate) fn set_row_col_after_sequence(&mut self) -> Result<(), Error> {
        let Some(header) = self.headers.get_mut(&self.current_field) else {
            return Err(serde::de::Error::custom(format!(
                "unknown field '{}', add it via Worksheet::add_serialize_headers()",
                self.current_field
            )));
        };

        header.row = self.current_row - 1;

        Ok(())
    }
}

pub(crate) struct SerializerHeader {
    row: RowNum,
    col: ColNum,
}

// -----------------------------------------------------------------------

impl Worksheet {
    /// TODO
    ///
    ///
    /// # Errors
    ///
    ///
    pub fn add_serialize_headers(
        &mut self,
        row: RowNum,
        col: ColNum,
        headers: &[&str],
    ) -> Result<&mut Worksheet, XlsxError> {
        // Check row and columns are in the allowed range.
        if !self.check_dimensions_only(row, col) {
            return Err(XlsxError::RowColumnLimitError);
        }

        let col_initial = col;

        for (col_offset, header) in headers.iter().enumerate() {
            let col = col_initial + col_offset as u16;

            let serializer_header = SerializerHeader { row, col };

            self.serializer_state
                .headers
                .insert((*header).to_string(), serializer_header);

            self.write(row, col, *header)?;
        }

        Ok(self)
    }

    pub(crate) fn serialize_to_worksheet_cell(
        &mut self,
        data: impl IntoExcelData,
    ) -> Result<&mut Worksheet, XlsxError> {
        let row = self.serializer_state.current_row;
        let col = self.serializer_state.current_col;

        self.write(row, col, data)?;

        Ok(self)
    }
}

// -----------------------------------------------------------------------

impl From<XlsxError> for serde::de::value::Error {
    fn from(e: XlsxError) -> serde::de::value::Error {
        serde::de::Error::custom(e.to_string())
    }
}

// -----------------------------------------------------------------------

pub fn to_worksheet_cells<T>(value: &T, serializer: &mut Worksheet) -> Result<(), Error>
where
    T: Serialize,
{
    value.serialize(serializer)?;
    Ok(())
}

// -----------------------------------------------------------------------

#[allow(unused_variables)]
impl<'a> ser::Serializer for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;
    type SerializeSeq = Self;
    type SerializeTuple = Self;
    type SerializeTupleStruct = Self;
    type SerializeTupleVariant = Self;
    type SerializeMap = Self;
    type SerializeStruct = Self;
    type SerializeStructVariant = Self;

    //
    // Serialize all the default number types that fit into Excel's f64 type.
    //
    fn serialize_bool(self, data: bool) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_i8(self, data: i8) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_u8(self, data: u8) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_i16(self, data: i16) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_u16(self, data: u16) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_i32(self, data: i32) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_u32(self, data: u32) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_f32(self, data: f32) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    fn serialize_f64(self, data: f64) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    //
    // Serialize i64/u64.
    //
    // Excel uses a f64 data type for all numbers. i64/u64 won't fit losslessly
    // into f64 but a large part of its range does (+/- 2^52). As such, we
    // convert the integers that will convert losslessly and raise an error for
    // anything outside that range.
    //
    #[allow(clippy::cast_precision_loss)]
    fn serialize_i64(self, data: i64) -> Result<(), Error> {
        if (MIN_LOSSLESS_I64_TO_F64..=MAX_LOSSLESS_I64_TO_F64).contains(&data) {
            self.serialize_f64(data as f64)
        } else {
            Err(serde::de::Error::custom(format!(
                "i64 value '{data}' does not fit into to Excel's f64 range"
            )))
        }
    }

    #[allow(clippy::cast_precision_loss)]
    fn serialize_u64(self, data: u64) -> Result<(), Error> {
        if (0..=MAX_LOSSLESS_U64_TO_F64).contains(&data) {
            self.serialize_f64(data as f64)
        } else {
            Err(serde::de::Error::custom(format!(
                "u64 value '{data}' does not fit into to Excel's f64 range"
            )))
        }
    }

    //
    // Serialize strings types.
    //
    fn serialize_str(self, data: &str) -> Result<(), Error> {
        self.serialize_to_worksheet_cell(data)?;
        Ok(())
    }

    // Serialize a char as a single-character string.
    fn serialize_char(self, data: char) -> Result<(), Error> {
        self.serialize_str(&data.to_string())
    }

    // Excel has no type equivalent to a byte array.
    fn serialize_bytes(self, data: &[u8]) -> Result<(), Error> {
        Err(serde::de::Error::custom(
            "byte array is not support by Excel. See the `rust_xlsxwriter` serialization docs for supported data types".to_string()
        ))
    }

    // Serialize Some(T).
    fn serialize_some<T>(self, data: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        data.serialize(self)
    }

    // Empty/None/Null values in Excel are ignored unless the cell has
    // formatting in which case they are handled as a "blank" cell. For all of
    // these cases we write an empty string and `rust_xlsxwriter` will handle it
    // correctly based on context.

    fn serialize_none(self) -> Result<(), Error> {
        self.serialize_str("")
    }

    fn serialize_unit(self) -> Result<(), Error> {
        self.serialize_none()
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<(), Error> {
        self.serialize_none()
    }

    // When serializing a unit variant (or any other kind of variant), formats
    // can choose whether to keep track of it by index or by name. Binary
    // formats typically use the index of the variant and human-readable formats
    // typically use the name.
    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
    ) -> Result<(), Error> {
        self.serialize_str(variant)
    }

    // As is done here, serializers are encouraged to treat newtype structs as
    // insignificant wrappers around the data they contain. TODO
    fn serialize_newtype_struct<T>(self, _name: &'static str, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(self)
    }

    // Note that newtype variant (and all of the other variant serialization
    // methods) refer exclusively to the "externally tagged" enum
    // representation. TODO
    //
    // Serialize this to JSON in externally tagged form as `{ NAME: VALUE }`.
    fn serialize_newtype_variant<T>(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        Ok(())
    }

    // Compound types.
    //
    // The only compound types that map into the Excel data model are
    // structs and array/vector types as fields in a struct.

    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq, Error> {
        Ok(self)
    }

    // TODO - not used.
    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, Error> {
        self.serialize_seq(Some(len))
    }

    // TODO - not used.
    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, Error> {
        self.serialize_seq(Some(len))
    }

    // TODO - not used?
    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleVariant, Error> {
        variant.serialize(&mut *self)?;
        Ok(self)
    }

    // The field/values of structs are treated as a map.
    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap, Error> {
        Ok(self)
    }

    // Structs are the main primary data type used to map data structures into
    // Excel.
    fn serialize_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeStruct, Error> {
        self.serialize_map(Some(len))
    }

    // TODO - not used.
    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStructVariant, Error> {
        variant.serialize(&mut *self)?;
        Ok(self)
    }
}

// The following 7 impls deal with the serialization of compound types.
// Currently we only support/use SerializeStruct and SerializeSeq.

impl<'a> ser::SerializeSeq for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    // Serialize a single element of the sequence.
    fn serialize_element<T>(&mut self, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        let ret = value.serialize(&mut **self);

        // Increment the row number for each element of the sequence.
        self.serializer_state.current_row += 1;

        ret
    }

    // Close the sequence.
    fn end(self) -> Result<(), Error> {
        self.serializer_state.set_row_col_after_sequence()?;

        Ok(())
    }
}

// Serialize tuple sequences.
impl<'a> ser::SerializeTuple for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    fn serialize_element<T>(&mut self, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        Ok(())
    }
}

// Serialize tuple struct sequences.
impl<'a> ser::SerializeTupleStruct for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T>(&mut self, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        Ok(())
    }
}

// Serialize tuple variant sequences.
impl<'a> ser::SerializeTupleVariant for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T>(&mut self, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        Ok(())
    }
}

// Serialize tuple map sequences.
impl<'a> ser::SerializeMap for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    fn serialize_key<T>(&mut self, key: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        key.serialize(&mut **self)
    }

    fn serialize_value<T>(&mut self, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        Ok(())
    }
}

// Structs are are the main sequence type used by `rust_xlsxwriter`.
impl<'a> ser::SerializeStruct for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T>(&mut self, key: &'static str, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        // Set the serializer (row, col) starting point for data serialization
        // by mapping the field name to a header name.
        self.serializer_state.set_row_col_for_field(key)?;

        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        Ok(())
    }
}

// Similar to `SerializeTupleVariant`, here the `end` method is responsible for
// closing both of the curly braces opened by `serialize_struct_variant`.
impl<'a> ser::SerializeStructVariant for &'a mut Worksheet {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T>(&mut self, key: &'static str, value: &T) -> Result<(), Error>
    where
        T: ?Sized + Serialize,
    {
        key.serialize(&mut **self)?;
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        Ok(())
    }
}
