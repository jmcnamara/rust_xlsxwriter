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

use crate::{ExcelDateTime, Formula, IntoExcelDateTime, XlsxError};
use std::fmt;

// -----------------------------------------------------------------------
// DataValidation
// -----------------------------------------------------------------------

/// The `DataValidation` struct represents a Cell conditional format.
///
/// TODO
///
#[derive(Clone)]
pub struct DataValidation {
    pub(crate) validation_type: DataValidationType,
    pub(crate) rule: DataValidationRuleInternal,
    pub(crate) ignore_blank: bool,
    pub(crate) show_input_message: bool,
    pub(crate) show_error_message: bool,
    pub(crate) multi_range: String,
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
            validation_type: DataValidationType::Any,
            rule: DataValidationRuleInternal::EqualTo(String::new()),
            ignore_blank: true,
            show_input_message: true,
            show_error_message: true,
            multi_range: String::new(),
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
    pub fn allow_any_value(mut self) -> DataValidation {
        self.rule = DataValidationRuleInternal::EqualTo(String::new());
        self.validation_type = DataValidationType::Any;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_whole_number(mut self, rule: DataValidationRule<i32>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Whole;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_whole_number_formula(
        mut self,
        rule: DataValidationRule<Formula>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Whole;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_decimal_number(mut self, rule: DataValidationRule<f64>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Decimal;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_decimal_number_formula(
        mut self,
        rule: DataValidationRule<Formula>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Decimal;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    /// # Errors
    ///
    pub fn allow_list_strings(
        mut self,
        list: &[impl AsRef<str>],
    ) -> Result<DataValidation, XlsxError> {
        let joined_list = list
            .iter()
            .map(|s| s.as_ref().to_string().replace('"', "\"\""))
            .collect::<Vec<String>>()
            .join(",");

        let length = joined_list.chars().count();
        if length > 255 {
            return Err(XlsxError::DataValidationError(
                format!("Validation list length '{length}' including commas is greater than Excel's limit of 255 characters: {joined_list}")
            ));
        }

        let joined_list = format!("\"{joined_list}\"");

        self.rule = DataValidationRuleInternal::ListSource(joined_list);
        self.validation_type = DataValidationType::List;
        Ok(self)
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_list_formula(mut self, rule: Formula) -> DataValidation {
        let formula = rule.expand_formula(true).to_string();
        self.rule = DataValidationRuleInternal::ListSource(formula);
        self.validation_type = DataValidationType::List;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_date(
        mut self,
        rule: DataValidationRule<impl IntoExcelDateTime + IntoDataValidationValue>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Date;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_date_formula(mut self, rule: DataValidationRule<Formula>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Date;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_time(
        mut self,
        rule: DataValidationRule<impl IntoExcelDateTime + IntoDataValidationValue>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Time;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_time_formula(mut self, rule: DataValidationRule<Formula>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Time;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_text_length(mut self, rule: DataValidationRule<u32>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::TextLength;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_text_length_formula(
        mut self,
        rule: DataValidationRule<Formula>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::TextLength;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn allow_custom_formula(mut self, rule: Formula) -> DataValidation {
        let formula = rule.expand_formula(true).to_string();
        self.rule = DataValidationRuleInternal::CustomFormula(formula);
        self.validation_type = DataValidationType::Custom;
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
    pub fn set_input_title(mut self, text: impl Into<String>) -> DataValidation {
        let text = text.into();
        let length = text.chars().count();

        if length > 32 {
            eprintln!(
                "Validation title length '{length}' greater than Excel's limit of 32 characters."
            );
            return self;
        }

        self.input_title = text;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_input_message(mut self, text: impl Into<String>) -> DataValidation {
        let text = text.into();
        let length = text.chars().count();

        if length > 255 {
            eprintln!(
                "Validation message length '{length}' greater than Excel's limit of 255 characters."
            );
            return self;
        }

        self.input_message = text;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_error_title(mut self, text: impl Into<String>) -> DataValidation {
        let text = text.into();
        let length = text.chars().count();

        if length > 32 {
            eprintln!(
                "Validation title length '{length}' greater than Excel's limit of 32 characters."
            );
            return self;
        }

        self.error_title = text;
        self
    }

    /// Set the TODO
    ///
    /// TODO
    ///
    pub fn set_error_message(mut self, text: impl Into<String>) -> DataValidation {
        let text = text.into();
        let length = text.chars().count();

        if length > 255 {
            eprintln!(
                "Validation message length '{length}' greater than Excel's limit of 255 characters."
            );
            return self;
        }

        self.error_message = text;
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

    /// Set an additional multi-cell range for the conditional format.
    ///
    /// The `set_multi_range()` method is used to extend a conditional
    /// format over non-contiguous ranges like `"B3:D6 I3:K6 B9:D12
    /// I9:K12"`.
    ///
    /// See [Selecting a non-contiguous
    /// range](crate::conditional_format#selecting-a-non-contiguous-range)
    /// for more information.
    ///
    /// # Parameters
    ///
    /// * `range` - A string like type representing an Excel range.
    ///
    ///   Note, you can use an Excel range like `"$B$3:$D$6,$I$3:$K$6"` or
    ///   omit the `$` anchors and replace the commas with spaces to have a
    ///   clearer range like `"B3:D6 I3:K6"`. The documentation and examples
    ///   use the latter format for clarity but it you are copying and
    ///   pasting from Excel you can use the first format.
    ///
    ///   Note, if the range is invalid then Excel will omit it silently.
    ///
    pub fn set_multi_range(mut self, range: impl Into<String>) -> DataValidation {
        self.multi_range = range.into().replace('$', "").replace(',', " ");
        self
    }

    // The "Any" validation type should be ignored if it doesn't have any input
    // or error titles or messages. This is the same rule as Excel.
    pub(crate) fn is_invalid_any(&mut self) -> bool {
        self.validation_type == DataValidationType::Any
            && self.input_title.is_empty()
            && self.input_message.is_empty()
            && self.error_title.is_empty()
            && self.error_message.is_empty()
    }
}

/// Trait to map rust types into data validation types
///
/// The `IntoDataValidationValue` trait is used to map Rust types like
/// strings, numbers, dates, times and formulas into a generic type that can be
/// used to replicate Excel data types used in Data Validation. TODO
///
pub trait IntoDataValidationValue {
    /// Function to turn types into a TODO enum value.
    fn to_string_value(&self) -> String;
}

impl IntoDataValidationValue for i32 {
    fn to_string_value(&self) -> String {
        self.to_string()
    }
}

impl IntoDataValidationValue for u32 {
    fn to_string_value(&self) -> String {
        self.to_string()
    }
}

impl IntoDataValidationValue for f64 {
    fn to_string_value(&self) -> String {
        self.to_string()
    }
}

impl IntoDataValidationValue for Formula {
    fn to_string_value(&self) -> String {
        self.expand_formula(true).to_string()
    }
}

impl IntoDataValidationValue for ExcelDateTime {
    fn to_string_value(&self) -> String {
        self.to_excel().to_string()
    }
}

impl IntoDataValidationValue for &ExcelDateTime {
    fn to_string_value(&self) -> String {
        self.to_excel().to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for NaiveDateTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_datetime_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for &NaiveDateTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_datetime_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for NaiveDate {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_date_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for &NaiveDate {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_date_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for NaiveTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_time_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for &NaiveTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_time_to_excel(self).to_string()
    }
}

//#[cfg(feature = "chrono")]
//#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
//data_validation_value_from_type!(&NaiveDate & NaiveDateTime & NaiveTime);

// -----------------------------------------------------------------------
// DataValidationType
// -----------------------------------------------------------------------

/// The `DataValidationType` enum defines TODO
///
///
#[derive(Clone, Eq, PartialEq)]
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

    /// TODO
    Custom,

    /// TODO
    List,

    /// TODO
    Any,
}

impl fmt::Display for DataValidationType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Any => write!(f, "any"),
            Self::Date => write!(f, "date"),
            Self::List => write!(f, "list"),
            Self::Time => write!(f, "time"),
            Self::Whole => write!(f, "whole"),
            Self::Custom => write!(f, "custom"),
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
    /// TODO.
    EqualTo(T),

    /// TODO.
    NotEqualTo(T),

    /// TODO.
    GreaterThan(T),

    /// TODO.
    GreaterThanOrEqualTo(T),

    /// TODO.
    LessThan(T),

    /// TODO.
    LessThanOrEqualTo(T),

    /// TODO.
    Between(T, T),

    /// TODO.
    NotBetween(T, T),
}

impl<T: IntoDataValidationValue> DataValidationRule<T> {
    fn to_internal_rule(&self) -> DataValidationRuleInternal {
        match &self {
            DataValidationRule::EqualTo(value) => {
                DataValidationRuleInternal::EqualTo(value.to_string_value())
            }
            DataValidationRule::NotEqualTo(value) => {
                DataValidationRuleInternal::NotEqualTo(value.to_string_value())
            }
            DataValidationRule::GreaterThan(value) => {
                DataValidationRuleInternal::GreaterThan(value.to_string_value())
            }

            DataValidationRule::GreaterThanOrEqualTo(value) => {
                DataValidationRuleInternal::GreaterThanOrEqualTo(value.to_string_value())
            }
            DataValidationRule::LessThan(value) => {
                DataValidationRuleInternal::LessThan(value.to_string_value())
            }
            DataValidationRule::LessThanOrEqualTo(value) => {
                DataValidationRuleInternal::LessThanOrEqualTo(value.to_string_value())
            }
            DataValidationRule::Between(min, max) => {
                DataValidationRuleInternal::Between(min.to_string_value(), max.to_string_value())
            }
            DataValidationRule::NotBetween(min, max) => {
                DataValidationRuleInternal::NotBetween(min.to_string_value(), max.to_string_value())
            }
        }
    }
}

// -----------------------------------------------------------------------
// DataValidationRuleInternal
// -----------------------------------------------------------------------

// TODO
#[derive(Clone)]
pub(crate) enum DataValidationRuleInternal {
    EqualTo(String),

    NotEqualTo(String),

    GreaterThan(String),

    GreaterThanOrEqualTo(String),

    LessThan(String),

    LessThanOrEqualTo(String),

    Between(String, String),

    NotBetween(String, String),

    CustomFormula(String),

    ListSource(String),
}

impl fmt::Display for DataValidationRuleInternal {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::EqualTo(_) => write!(f, "equal"),
            Self::LessThan(_) => write!(f, "lessThan"),
            Self::Between(_, _) => write!(f, "between"),
            Self::ListSource(_) => write!(f, "list"),
            Self::NotEqualTo(_) => write!(f, "notEqual"),
            Self::GreaterThan(_) => write!(f, "greaterThan"),
            Self::CustomFormula(_) => write!(f, ""),
            Self::NotBetween(_, _) => write!(f, "notBetween"),
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
