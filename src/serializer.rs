// serializer - A serde serializer for use with `rust_xlsxwriter`.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use std::collections::HashMap;
use std::fmt::Display;

use crate::{ColNum, IntoExcelData, RowNum, Worksheet, XlsxError};
use serde::{ser, Serialize};

const MAX_LOSSLESS_U64_TO_F64: u64 = 2 << 52;
const MAX_LOSSLESS_I64_TO_F64: i64 = 2 << 52;
const MIN_LOSSLESS_I64_TO_F64: i64 = -MAX_LOSSLESS_I64_TO_F64;

// -----------------------------------------------------------------------
// Worksheet extensions to handle serialization.
// -----------------------------------------------------------------------

// The serialization Worksheet methods are added in this module to make it
// easier to isolate the feature specific code.
impl Worksheet {
    /// TODO - Add full documentation and examples.
    ///
    /// # Errors
    ///
    pub fn serialize<T>(&mut self, data: &T) -> Result<&mut Worksheet, XlsxError>
    where
        T: Serialize,
    {
        self.serialize_data_structure(data)?;

        Ok(self)
    }

    /// TODO - This will be replaced with a serialized version.
    ///
    ///
    /// # Errors
    ///
    ///
    pub fn write_serialize_headers(
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

    // Serialize the parent data structure to the worksheet.
    pub(crate) fn serialize_data_structure<T>(
        &mut self,
        data_structure: &T,
    ) -> Result<(), XlsxError>
    where
        T: Serialize,
    {
        data_structure.serialize(self)?;
        Ok(())
    }

    // Serialize individual data items to a worksheet cell.
    pub(crate) fn serialize_to_worksheet_cell(
        &mut self,
        data: impl IntoExcelData,
    ) -> Result<(), XlsxError> {
        let row = self.serializer_state.current_row;
        let col = self.serializer_state.current_col;

        self.write(row, col, data)?;

        Ok(())
    }
}

// -----------------------------------------------------------------------
// SerializerState, a struct to maintain row/colum state and other metadata
// between serialized writes. This avoids passing around cell location
// information in the serializer.
// -----------------------------------------------------------------------
pub(crate) struct SerializerState {
    headers: HashMap<String, SerializerHeader>,
    current_field: String,
    current_col: ColNum,
    current_row: RowNum,
}

impl SerializerState {
    // Create a new SerializerState struct.
    pub(crate) fn new() -> SerializerState {
        SerializerState {
            headers: HashMap::new(),
            current_field: String::new(),
            current_col: 0,
            current_row: 0,
        }
    }

    // Set the row/col data to use when writing a serialized value. The column
    // number comes from the initial serialization of the headers and the row is
    // incremented with each access to a field.
    pub(crate) fn set_row_col_for_field(&mut self, key: &'static str) -> Result<(), XlsxError> {
        let Some(header) = self.headers.get_mut(key) else {
            return Err(XlsxError::ParameterError(format!(
                "unknown field '{key}', add it via Worksheet::add_serialize_headers()"
            )));
        };

        // Increment the row number for the next worksheet.write().
        header.row += 1;

        // Set the "current" cell values used to write the serialized data.
        self.current_col = header.col;
        self.current_row = header.row;
        self.current_field = key.to_string();

        Ok(())
    }

    // Store the last row position after writing a vec/array sequence.
    pub(crate) fn set_row_col_after_sequence(&mut self) -> Result<(), XlsxError> {
        let Some(header) = self.headers.get_mut(&self.current_field) else {
            return Err(XlsxError::ParameterError(format!(
                "unknown field '{}', add it via Worksheet::add_serialize_headers()",
                self.current_field
            )));
        };

        // Store the row position for the field.
        header.row = self.current_row - 1;

        Ok(())
    }
}

// -----------------------------------------------------------------------
// SerializerHeader. A struct used to store header/field position and metadata
// for individual fields.
// -----------------------------------------------------------------------
pub(crate) struct SerializerHeader {
    row: RowNum,
    col: ColNum,
}

// -----------------------------------------------------------------------
// Implementation of the serde::ser::Error Trait for XlsxWriter to allow us to
// use a single error type for both serialization and writing errors.
// -----------------------------------------------------------------------
impl serde::ser::Error for XlsxError {
    fn custom<T: Display>(msg: T) -> Self {
        XlsxError::SerdeError(msg.to_string())
    }
}

// -----------------------------------------------------------------------
// Serializer. This is the implementation of the Serializer trait to serialized
// a serde derived struct to an Excel worksheet.
// -----------------------------------------------------------------------
#[allow(unused_variables)]
impl<'a> ser::Serializer for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;
    type SerializeSeq = Self;
    type SerializeTuple = Self;
    type SerializeTupleStruct = Self;
    type SerializeTupleVariant = Self;
    type SerializeMap = Self;
    type SerializeStruct = Self;
    type SerializeStructVariant = Self;

    // Serialize all the default number types that fit into Excel's f64 type.
    fn serialize_bool(self, data: bool) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_i8(self, data: i8) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_u8(self, data: u8) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_i16(self, data: i16) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_u16(self, data: u16) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_i32(self, data: i32) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_u32(self, data: u32) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_f32(self, data: f32) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    fn serialize_f64(self, data: f64) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    // Serialize i64/u64.
    //
    // Excel uses a f64 data type for all numbers. i64/u64 won't fit losslessly
    // into f64 but a large part of its range does (+/- 2^52). As a compromise,
    // we convert the integers that will convert losslessly and raise an error
    // for anything outside that range.
    #[allow(clippy::cast_precision_loss)]
    fn serialize_i64(self, data: i64) -> Result<(), XlsxError> {
        if (MIN_LOSSLESS_I64_TO_F64..=MAX_LOSSLESS_I64_TO_F64).contains(&data) {
            self.serialize_f64(data as f64)
        } else {
            Err(XlsxError::SerdeError(format!(
                "i64 value '{data}' does not fit into to Excel's f64 range"
            )))
        }
    }

    #[allow(clippy::cast_precision_loss)]
    fn serialize_u64(self, data: u64) -> Result<(), XlsxError> {
        if (0..=MAX_LOSSLESS_U64_TO_F64).contains(&data) {
            self.serialize_f64(data as f64)
        } else {
            Err(XlsxError::SerdeError(format!(
                "u64 value '{data}' does not fit into to Excel's f64 range"
            )))
        }
    }

    // Serialize strings types.
    fn serialize_str(self, data: &str) -> Result<(), XlsxError> {
        self.serialize_to_worksheet_cell(data)
    }

    // Excel doesn't have a character type. Serialize a char as a
    // single-character string.
    fn serialize_char(self, data: char) -> Result<(), XlsxError> {
        self.serialize_str(&data.to_string())
    }

    // Excel doesn't have a type equivalent to a byte array.
    fn serialize_bytes(self, data: &[u8]) -> Result<(), XlsxError> {
        Err(XlsxError::SerdeError(
            "byte array is not support by Excel. See the `rust_xlsxwriter` serialization docs for supported data types".to_string()
        ))
    }

    // Serialize Some(T) values.
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

    fn serialize_none(self) -> Result<(), XlsxError> {
        self.serialize_str("")
    }

    fn serialize_unit(self) -> Result<(), XlsxError> {
        self.serialize_none()
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<(), XlsxError> {
        self.serialize_none()
    }

    // Excel doesn't have an equivalent for the structure so we ignore it.
    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
    ) -> Result<(), XlsxError> {
        Ok(())
    }

    // Try to handle this as a single value.
    fn serialize_newtype_struct<T>(self, _name: &'static str, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        value.serialize(self)
    }

    // Excel doesn't have an equivalent for the structure so we ignore it.
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

    // Compound types.
    //
    // The only compound types that map into the Excel data model are
    // structs and array/vector types as fields in a struct.

    // Structs are the main primary data type used to map data structures into
    // Excel.
    fn serialize_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeStruct, XlsxError> {
        self.serialize_map(Some(len))
    }

    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq, XlsxError> {
        Ok(self)
    }

    // Not used.
    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple, XlsxError> {
        self.serialize_seq(Some(len))
    }

    // Not used.
    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct, XlsxError> {
        self.serialize_seq(Some(len))
    }

    // Not used.
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
    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap, XlsxError> {
        Ok(self)
    }

    // Not used.
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
impl<'a> ser::SerializeStruct for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    fn serialize_field<T>(&mut self, key: &'static str, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        // Set the serializer (row, col) starting point for data serialization
        // by mapping the field name to a header name.
        self.serializer_state.set_row_col_for_field(key)?;

        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), XlsxError> {
        Ok(())
    }
}

// We also serialize sequences to map vectors/arrays to Excel.
impl<'a> ser::SerializeSeq for &'a mut Worksheet {
    type Ok = ();
    type Error = XlsxError;

    // Serialize a single element of the sequence.
    fn serialize_element<T>(&mut self, value: &T) -> Result<(), XlsxError>
    where
        T: ?Sized + Serialize,
    {
        let ret = value.serialize(&mut **self);

        // Increment the row number for each element of the sequence.
        self.serializer_state.current_row += 1;

        ret
    }

    // Close the sequence.
    fn end(self) -> Result<(), XlsxError> {
        self.serializer_state.set_row_col_after_sequence()
    }
}

// Serialize tuple sequences.
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
