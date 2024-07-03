// data_validation - A module to represent Excel data validations.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! # Working with Data Validation
//!
//! TODO

#![warn(missing_docs)]

mod tests;

#[cfg(feature = "chrono")]
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

use crate::{ExcelDateTime, Formula, XlsxError};
use std::{borrow::Cow, fmt};

// -----------------------------------------------------------------------
// DataValidation
// -----------------------------------------------------------------------

/// The `DataValidation` struct represents a Cell conditional format.
///
/// TODO
///
#[derive(Clone)]
pub struct DataValidation {
    pub(crate) validation_type: Option<DataValidationType>,
    pub(crate) rule: Option<DataValidationRule<DataValidationValue>>,
    pub(crate) ignore_blank: bool,
    pub(crate) show_input_message: bool,
    pub(crate) show_error_message: bool,
    pub(crate) input_title: String,
    pub(crate) error_title: String,
    pub(crate) input_message: String,
    pub(crate) error_message: String,
    pub(crate) error_style: DataValidationErrorStyle,
}

impl DataValidation {
    /// Create a new Cell conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> DataValidation {
        DataValidation {
            validation_type: None,
            rule: None,
            ignore_blank: true,
            show_input_message: true,
            show_error_message: true,
            input_title: String::new(),
            error_title: String::new(),
            input_message: String::new(),
            error_message: String::new(),
            error_style: DataValidationErrorStyle::Stop,
        }
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_type(mut self, validation_type: DataValidationType) -> DataValidation {
        self.validation_type = Some(validation_type);
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_rule<T>(mut self, rule: DataValidationRule<T>) -> DataValidation
    where
        T: IntoDataValidationValue,
    {
        // Change from a generic type to a concrete DataValidationValue type.
        let rule = match rule {
            DataValidationRule::EqualTo(value) => DataValidationRule::EqualTo(value.new_value()),
            DataValidationRule::NotEqualTo(value) => {
                DataValidationRule::NotEqualTo(value.new_value())
            }
            DataValidationRule::GreaterThan(value) => {
                DataValidationRule::GreaterThan(value.new_value())
            }
            DataValidationRule::GreaterThanOrEqualTo(value) => {
                DataValidationRule::GreaterThanOrEqualTo(value.new_value())
            }
            DataValidationRule::LessThan(value) => DataValidationRule::LessThan(value.new_value()),
            DataValidationRule::LessThanOrEqualTo(value) => {
                DataValidationRule::LessThanOrEqualTo(value.new_value())
            }
            DataValidationRule::Between(min, max) => {
                DataValidationRule::Between(min.new_value(), max.new_value())
            }
            DataValidationRule::NotBetween(min, max) => {
                DataValidationRule::NotBetween(min.new_value(), max.new_value())
            }
        };

        self.rule = Some(rule);
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn ignore_blank(mut self, enable: bool) -> DataValidation {
        self.ignore_blank = enable;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn show_input_message(mut self, enable: bool) -> DataValidation {
        self.show_input_message = enable;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn show_error_message(mut self, enable: bool) -> DataValidation {
        self.show_error_message = enable;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_input_title(mut self, title: impl Into<String>) -> DataValidation {
        // TODO add string length check.
        self.input_title = title.into();
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_input_message(mut self, title: impl Into<String>) -> DataValidation {
        // TODO add string length check.
        self.input_message = title.into();
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_error_title(mut self, title: impl Into<String>) -> DataValidation {
        // TODO add string length check.
        self.error_title = title.into();
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_error_message(mut self, title: impl Into<String>) -> DataValidation {
        // TODO add string length check.
        self.error_message = title.into();
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_error_style(mut self, error_style: DataValidationErrorStyle) -> DataValidation {
        self.error_style = error_style;
        self
    }

    // Validate the conditional format.
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        if self.rule.is_none() {
            return Err(XlsxError::DataValidationError(
                "DataValidation rule must be set".to_string(),
            ));
        }

        Ok(())
    }
}

// -----------------------------------------------------------------------
// DataValidationValue
// -----------------------------------------------------------------------

/// The `DataValidationValue` struct represents conditional format value
/// types. TODO
///
/// Excel supports various types when specifying values in a conditional format
/// such as numbers, strings, dates, times and cell references.
/// `DataValidationValue` is used to support a similar generic interface to
/// conditional format values. It supports:
///
/// - Numbers: Any Rust number that can convert [`Into`] [`f64`].
/// - Strings: Any Rust string type that can convert into String such as
///   [`&str`], [`String`], `&String` and `Cow<'_, str>`.
/// - Dates/times: [`ExcelDateTime`] values and if the `chrono` feature is
///   enabled [`chrono::NaiveDateTime`], [`chrono::NaiveDate`] and
///   [`chrono::NaiveTime`].
/// - Cell ranges: Use [`Formula`] in order to distinguish from strings. For
///   example `Formula::new(=A1)`.
///
/// [`chrono::NaiveDate`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
/// [`chrono::NaiveTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
/// [`chrono::NaiveDateTime`]: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
///
#[derive(Clone)]
pub struct DataValidationValue {
    pub(crate) value: String,
    pub(crate) is_string: bool,
}

impl DataValidationValue {
    pub(crate) fn new_from_string(value: impl Into<String>) -> DataValidationValue {
        DataValidationValue {
            value: value.into(),
            is_string: false,
        }
    }
}

// From/Into traits for DataValidationValue.
macro_rules! data_validation_value_from_string {
    ($($t:ty)*) => ($(
        impl From<$t> for DataValidationValue {
            fn from(value: $t) -> DataValidationValue {
                let mut value = DataValidationValue::new_from_string(value);
                value.is_string = true;
                value
            }
        }
    )*)
}
data_validation_value_from_string!(&str &String String Cow<'_, str>);

macro_rules! data_validation_value_from_number {
    ($($t:ty)*) => ($(
        impl From<$t> for DataValidationValue {
            fn from(value: $t) -> DataValidationValue {
                DataValidationValue::new_from_string(value.to_string())
            }
        }
    )*)
}
data_validation_value_from_number!(u8 i8 u16 i16 u32 i32 f32 f64);

impl From<Formula> for DataValidationValue {
    fn from(value: Formula) -> DataValidationValue {
        DataValidationValue::new_from_string(value.expand_formula(true))
    }
}

impl From<ExcelDateTime> for DataValidationValue {
    fn from(value: ExcelDateTime) -> DataValidationValue {
        let value = value.to_excel().to_string();
        DataValidationValue::new_from_string(value)
    }
}

impl From<&ExcelDateTime> for DataValidationValue {
    fn from(value: &ExcelDateTime) -> DataValidationValue {
        let value = value.to_excel().to_string();
        DataValidationValue::new_from_string(value)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl From<&NaiveDate> for DataValidationValue {
    fn from(value: &NaiveDate) -> DataValidationValue {
        let value = ExcelDateTime::chrono_date_to_excel(value).to_string();
        DataValidationValue::new_from_string(value)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl From<&NaiveDateTime> for DataValidationValue {
    fn from(value: &NaiveDateTime) -> DataValidationValue {
        let value = ExcelDateTime::chrono_datetime_to_excel(value).to_string();
        DataValidationValue::new_from_string(value)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl From<&NaiveTime> for DataValidationValue {
    fn from(value: &NaiveTime) -> DataValidationValue {
        let value = ExcelDateTime::chrono_time_to_excel(value).to_string();
        DataValidationValue::new_from_string(value)
    }
}

/// Trait to map rust types into an [`DataValidationValue`] value.
///
/// The `IntoDataValidationValue` trait is used to map Rust types like
/// strings, numbers, dates, times and formulas into a generic type that can be
/// used to replicate Excel data types used in Data Validation.
///
/// See [`DataValidationValue`] for more information.
///
pub trait IntoDataValidationValue {
    /// Function to turn types into a [`DataValidationValue`] enum value.
    fn new_value(self) -> DataValidationValue;
}

impl IntoDataValidationValue for DataValidationValue {
    fn new_value(self) -> DataValidationValue {
        self.clone()
    }
}

macro_rules! data_validation_value_from_type {
    ($($t:ty)*) => ($(
        impl IntoDataValidationValue for $t {
            fn new_value(self) -> DataValidationValue {
                self.into()
            }
        }
    )*)
}

data_validation_value_from_type!(
    &str &String String Cow<'_, str>
    u8 i8 u16 i16 u32 i32 f32 f64
    Formula
    ExcelDateTime &ExcelDateTime
);

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
data_validation_value_from_type!(&NaiveDate & NaiveDateTime & NaiveTime);

// -----------------------------------------------------------------------
// DataValidationType
// -----------------------------------------------------------------------

/// The `DataValidationType` enum defines TODO
///
///
#[derive(Clone)]
pub enum DataValidationType {
    /// TODO
    Whole,

    /// TODO
    Decimal,

    /// TODO
    Date,

    /// TODO
    Time,

    /// TODO
    TextLength,
}

impl fmt::Display for DataValidationType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Date => write!(f, "date"),
            Self::Time => write!(f, "time"),
            Self::Whole => write!(f, "whole"),
            Self::Decimal => write!(f, "decimal"),
            Self::TextLength => write!(f, "textLength"),
        }
    }
}

// -----------------------------------------------------------------------
// DataValidationRule
// -----------------------------------------------------------------------

/// The `DataValidationRule` enum defines the conditional format rule for
/// [`DataValidation`].
///
///
#[derive(Clone)]
pub enum DataValidationRule<T: IntoDataValidationValue> {
    /// Show the conditional format for cells that are equal to the target value.
    EqualTo(T),

    /// Show the conditional format for cells that are not equal to the target value.
    NotEqualTo(T),

    /// Show the conditional format for cells that are greater than the target value.
    GreaterThan(T),

    /// Show the conditional format for cells that are greater than or equal to the target value.
    GreaterThanOrEqualTo(T),

    /// Show the conditional format for cells that are less than the target value.
    LessThan(T),

    /// Show the conditional format for cells that are less than or equal to the target value.
    LessThanOrEqualTo(T),

    /// Show the conditional format for cells that are between the target values.
    Between(T, T),

    /// Show the conditional format for cells that are not between the target values.
    NotBetween(T, T),
}

impl<T: IntoDataValidationValue> fmt::Display for DataValidationRule<T> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::EqualTo(_) => write!(f, "equal"),
            Self::Between(_, _) => write!(f, "between"),
            Self::LessThan(_) => write!(f, "lessThan"),
            Self::NotEqualTo(_) => write!(f, "notEqual"),
            Self::NotBetween(_, _) => write!(f, "notBetween"),
            Self::GreaterThan(_) => write!(f, "greaterThan"),
            Self::LessThanOrEqualTo(_) => write!(f, "lessThanOrEqual"),
            Self::GreaterThanOrEqualTo(_) => write!(f, "greaterThanOrEqual"),
        }
    }
}

// -----------------------------------------------------------------------
// DataValidationErrorStyle
// -----------------------------------------------------------------------

/// The `DataValidationErrorStyle` enum defines TODO
///
///
#[derive(Clone)]
pub enum DataValidationErrorStyle {
    /// TODO
    Stop,

    /// TODO
    Warning,

    /// TODO
    Information,
}

impl fmt::Display for DataValidationErrorStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Stop => write!(f, "stop"),
            Self::Warning => write!(f, "warning"),
            Self::Information => write!(f, "information"),
        }
    }
}
