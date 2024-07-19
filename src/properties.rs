// properties - A module for representing document properties.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

#[cfg(feature = "chrono")]
use chrono::{DateTime, Utc};

use crate::ExcelDateTime;

/// The `DocProperties` struct is used to create an object to represent document
/// metadata properties.
///
/// The `DocProperties` struct is used to create an object to represent various
/// document properties for an Excel file such as the Author's name or the
/// Creation Date.
///
/// <img src="https://rustxlsxwriter.github.io/images/app_doc_properties.png">
///
/// Document Properties can be set for the "Summary" section and also for the
/// "Custom" section of the Excel document properties. See the examples below.
///
/// The `DocProperties` struct is used in conjunction with the
/// [`Workbook::set_properties()`](crate::Workbook::set_properties) method.
///
/// # Examples
///
/// An example of setting workbook document properties for a file created using
/// the `rust_xlsxwriter` library. This creates the file used to generate the
/// above image.
///
/// ```
/// # // This code is available in examples/app_doc_properties.rs
/// #
/// use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     let properties = DocProperties::new()
///         .set_title("This is an example spreadsheet")
///         .set_subject("That demonstrates document properties")
///         .set_author("A. Rust User")
///         .set_manager("J. Alfred Prufrock")
///         .set_company("Rust Solutions Inc")
///         .set_category("Sample spreadsheets")
///         .set_keywords("Sample, Example, DocProperties")
///         .set_comment("Created with Rust and rust_xlsxwriter");
///
///     workbook.set_properties(&properties);
///
///     let worksheet = workbook.add_worksheet();
///
///     worksheet.set_column_width(0, 30)?;
///     worksheet.write_string(0, 0, "See File -> Info -> Properties")?;
///
///     workbook.save("doc_properties.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// An example of setting custom/user defined workbook document properties.
///
/// ```
/// # // This code is available in examples/doc_properties_custom.rs
/// #
/// use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     let properties = DocProperties::new()
///         .set_custom_property("Checked by", "Admin")
///         .set_custom_property("Cross check", true)
///         .set_custom_property("Department", "Finance")
///         .set_custom_property("Document number", 55301);
///
///     workbook.set_properties(&properties);
///
///     workbook.save("properties.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/doc_properties_custom.png">
///
///
/// # Setting the Sensitivity Label for a file
///
/// Sensitivity Labels are a property that can be added to an Office 365
/// document to indicate that it is compliant with a company's information
/// protection policies. Sensitivity Labels have designations like
/// "Confidential", "Internal use only", or "Public" depending on the policies
/// implemented by the company. They are generally only enabled for enterprise
/// versions of Office.
///
/// See the following Microsoft documentation on how to [Apply sensitivity
/// labels to your files and email].
///
/// Sensitivity Labels are generally stored as custom document properties so
/// they can be enabled using [`DocProperties::set_custom_property()`]. However,
/// since the metadata differs from company to company you will need to extract
/// some of the required metadata from sample files.
///
/// [`DocProperties::set_custom_property()`]:
///     crate::DocProperties::set_custom_property
///
/// The first step is to create a new file in Excel and set a non-encrypted
/// sensitivity label. Then unzip the file by changing the extension from
/// `.xlsx` to `.zip` or by using a command line utility like this:
///
/// ```bash
/// $ unzip myfile.xlsx -d myfile
/// Archive:  myfile.xlsx
///   inflating: myfile/[Content_Types].xml
///   inflating: myfile/docProps/app.xml
///   inflating: myfile/docProps/custom.xml
///   inflating: myfile/docProps/core.xml
///   inflating: myfile/_rels/.rels
///   inflating: myfile/xl/workbook.xml
///   inflating: myfile/xl/worksheets/sheet1.xml
///   inflating: myfile/xl/styles.xml
///   inflating: myfile/xl/theme/theme1.xml
///   inflating: myfile/xl/_rels/workbook.xml.rels
/// ```
///
/// Then examine the `docProps/custom.xml` file from the unzipped xlsx file. The
/// file doesn't contain newlines so it is best to view it in an editor that can
/// handle XML or use a commandline utility like libxml’s [xmllint] to format
/// the XML for clarity:
///
/// [xmllint]: http://xmlsoft.org/xmllint.html
///
/// ```xml
/// $ xmllint --format myfile/docProps/custom.xml
/// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
/// <Properties
///     xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
///     xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="2"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Enabled">
///     <vt:lpwstr>true</vt:lpwstr>
///   </property>
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="3"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SetDate">
///     <vt:lpwstr>2024-01-01T12:00:00Z</vt:lpwstr>
///   </property>
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="4"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Method">
///     <vt:lpwstr>Privileged</vt:lpwstr>
///   </property>
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="5"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Name">
///     <vt:lpwstr>Confidential</vt:lpwstr>
///   </property>
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="6"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SiteId">
///     <vt:lpwstr>cb46c030-1825-4e81-a295-151c039dbf02</vt:lpwstr>
///   </property>
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="7"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ActionId">
///     <vt:lpwstr>88124cf5-1340-457d-90e1-0000a9427c99</vt:lpwstr>
///   </property>
///   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
///             pid="8"
///             name="MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ContentBits">
///     <vt:lpwstr>2</vt:lpwstr>
///   </property>
/// </Properties>
/// ```
///
/// The MSIP (Microsoft Information Protection) labels in the `name` attributes
/// contain a GUID that is unique to each company. The `SiteId` field will also
/// be unique to your company/location. The meaning of each of these fields is
/// explained in the the following Microsoft document on [Microsoft Information
/// Protection SDK - Metadata]. Once you have identified the necessary metadata
/// you can add it to a new document as shown below.
///
/// [Microsoft Information Protection SDK - Metadata]:
///     https://learn.microsoft.com/en-us/information-protection/develop/concept-mip-metadata
///
///
/// ```
/// # // This code is available in examples/app_sensitivity_label.rs
/// #
/// use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     // Metadata extracted from a company specific file.
///     let site_id = "cb46c030-1825-4e81-a295-151c039dbf02";
///     let action_id = "88124cf5-1340-457d-90e1-0000a9427c99";
///     let company_guid = "2096f6a2-d2f7-48be-b329-b73aaa526e5d";
///
///     // Add the document properties. Note that these should all be in text format.
///     let properties = DocProperties::new()
///         .set_custom_property(format!("MSIP_Label_{company_guid}_Method"), "Privileged")
///         .set_custom_property(format!("MSIP_Label_{company_guid}_Name"), "Confidential")
///         .set_custom_property(format!("MSIP_Label_{company_guid}_SiteId"), site_id)
///         .set_custom_property(format!("MSIP_Label_{company_guid}_ActionId"), action_id)
///         .set_custom_property(format!("MSIP_Label_{company_guid}_ContentBits"), "2");
///
///     workbook.set_properties(&properties);
///
///     workbook.save("sensitivity_label.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/app_sensitivity_label.png">
///
/// Note, some sensitivity labels require that the document is encrypted. In
/// order to extract the required metadata you will need to unencrypt the file
/// which may remove the sensitivity label. In that case you may need to use a
/// third party tool such as [msoffice-crypt].
///
/// [msoffice-crypt]: https://github.com/herumi/msoffice
///
/// [Apply sensitivity labels to your files and email]:
///     https://support.microsoft.com/en-us/office/apply-sensitivity-labels-to-your-files-and-email-2f96e7cd-d5a4-403b-8bd7-4cc636bae0f9
///
///
/// # Checksum of a saved file
///
/// A common issue that occurs with `rust_xlsxwriter`, but also with Excel, is
/// that running the same program twice doesn't generate the same file, byte for
/// byte. This can cause issues with applications that do checksumming for
/// testing purposes.
///
/// For example consider the following simple `rust_xlsxwriter` program:
///
/// ```
/// # // This code is available in examples/doc_properties_checksum1.rs
/// #
/// use rust_xlsxwriter::{Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     worksheet.write_string(0, 0, "Hello")?;
///
///     workbook.save("properties.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// If we run this several times, with a small delay, we will get different
/// checksums as shown below:
///
/// ```bash
/// $ cargo run --example doc_properties_checksum1
///
/// $ sum properties.xlsx
/// 62457 6 properties.xlsx
///
/// $ sleep 2
///
/// $ cargo run --example doc_properties_checksum1
///
/// $ sum properties.xlsx
/// 56692 6 properties.xlsx # Different to previous.
/// ```
///
/// This is due to a file creation datetime that is included in the file and
/// which changes each time a new file is created.
///
/// The relevant section of the `docProps/core.xml` sub-file in the xlsx format
/// looks like this:
///
///
/// ```xml
/// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
/// <cp:coreProperties>
///   <dc:creator/>
///   <cp:lastModifiedBy/>
///   <dcterms:created xsi:type="dcterms:W3CDTF">2023-01-08T00:23:58Z</dcterms:created>
///   <dcterms:modified xsi:type="dcterms:W3CDTF">2023-01-08T00:23:58Z</dcterms:modified>
/// </cp:coreProperties>
/// ```
///
/// If required this can be avoided by setting a constant creation date in the
/// document properties metadata:
///
///
/// ```
/// # // This code is available in examples/doc_properties_checksum2.rs
/// #
/// use rust_xlsxwriter::{DocProperties, ExcelDateTime, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     // Create a file creation date for the file.
///     let date = ExcelDateTime::from_ymd(2023, 1, 1)?;
///
///     // Add it to the document metadata.
///     let properties = DocProperties::new().set_creation_datetime(&date);
///     workbook.set_properties(&properties);
///
///     let worksheet = workbook.add_worksheet();
///     worksheet.write_string(0, 0, "Hello")?;
///
///     workbook.save("properties.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Then we will get the same checksum for the same output every time:
///
/// ```bash
/// $ cargo run --example doc_properties_checksum2
///
/// $ sum properties.xlsx
/// 8914 6 properties.xlsx
///
/// $ sleep 2
///
/// $ cargo run --example doc_properties_checksum2
///
/// $ sum properties.xlsx
/// 8914 6 properties.xlsx # Same as previous
/// ```
///
#[derive(Clone)]
pub struct DocProperties {
    pub(crate) author: String,
    pub(crate) title: String,
    pub(crate) comment: String,
    pub(crate) company: String,
    pub(crate) manager: String,
    pub(crate) status: String,
    pub(crate) subject: String,
    pub(crate) category: String,
    pub(crate) keywords: String,
    pub(crate) hyperlink_base: String,
    pub(crate) creation_time: String,
    pub(crate) custom_properties: Vec<CustomProperty>,
}

impl Default for DocProperties {
    fn default() -> Self {
        Self::new()
    }
}

impl DocProperties {
    /// Create a new `DocProperties` struct.
    pub fn new() -> DocProperties {
        DocProperties {
            title: String::new(),
            status: String::new(),
            author: String::new(),
            comment: String::new(),
            company: String::new(),
            manager: String::new(),
            subject: String::new(),
            category: String::new(),
            keywords: String::new(),
            hyperlink_base: String::new(),
            creation_time: ExcelDateTime::utc_now(),
            custom_properties: vec![],
        }
    }

    /// Set the Title field of the document properties.
    ///
    /// Set the "Title" field of the document properties to create a title for
    /// the document such as "Sales Report". See the example above.
    ///
    /// # Parameters
    ///
    /// - `title`: The title string property.
    ///
    pub fn set_title(mut self, title: impl Into<String>) -> DocProperties {
        self.title = title.into();

        self
    }

    /// Set the Subject field of the document properties.
    ///
    /// Set the "Subject" field of the document properties to indicate the
    /// subject matter. See the example above.
    ///
    /// # Parameters
    ///
    /// - `subject`: The subject string property.
    ///
    pub fn set_subject(mut self, subject: impl Into<String>) -> DocProperties {
        self.subject = subject.into();

        self
    }

    /// Set the Manager field of the document properties.
    ///
    /// Set the "Manager" field of the document properties. See the example
    /// above. See the example above.
    ///
    /// # Parameters
    ///
    /// - `manager`: The manager string property.
    ///
    pub fn set_manager(mut self, manager: impl Into<String>) -> DocProperties {
        self.manager = manager.into();

        self
    }

    /// Set the Company field of the document properties.
    ///
    /// Set the "Company" field of the document properties. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// - `company`: The company string property.
    ///
    pub fn set_company(mut self, company: impl Into<String>) -> DocProperties {
        self.company = company.into();

        self
    }

    /// Set the Category field of the document properties.
    ///
    /// Set the "Category" field of the document properties to indicate the
    /// category that the file belongs to. See the example above.
    ///
    /// # Parameters
    ///
    /// - `category`: The category string property.
    ///
    pub fn set_category(mut self, category: impl Into<String>) -> DocProperties {
        self.category = category.into();

        self
    }

    /// Set the Author field of the document properties.
    ///
    /// Set the "Author" field of the document properties. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// - `author`: The author string property.
    ///
    pub fn set_author(mut self, author: impl Into<String>) -> DocProperties {
        self.author = author.into();

        self
    }

    /// Set the Keywords field of the document properties.
    ///
    /// Set the "Keywords" field of the document properties. This can be one or
    /// more keywords that can be used in searches. See the example above.
    ///
    /// # Parameters
    ///
    /// - `keywords`: The keywords string property.
    ///
    pub fn set_keywords(mut self, keywords: impl Into<String>) -> DocProperties {
        self.keywords = keywords.into();

        self
    }

    /// Set the Comment field of the document properties.
    ///
    /// Set the "Comment" field of the document properties. This can be a
    /// general comment or summary that you want to add to the properties. See
    /// the example above.
    ///
    /// # Parameters
    ///
    /// - `comment`: The comment string property.
    ///
    pub fn set_comment(mut self, comment: impl Into<String>) -> DocProperties {
        self.comment = comment.into();

        self
    }

    /// Set the Status field of the document properties.
    ///
    /// Set the "Status" field of the document properties such as "Draft" or
    /// "Final".
    ///
    /// # Parameters
    ///
    /// - `status`: The status string property.
    ///
    pub fn set_status(mut self, status: impl Into<String>) -> DocProperties {
        self.status = status.into();

        self
    }

    /// Set the hyperlink base field of the document properties.
    ///
    /// Set the "Hyperlink base" field of the document properties to have a
    /// default base url.
    ///
    /// # Parameters
    ///
    /// - `hyperlink_base`: The hyperlink base string property.
    ///
    pub fn set_hyperlink_base(mut self, hyperlink_base: impl Into<String>) -> DocProperties {
        self.hyperlink_base = hyperlink_base.into();

        self
    }

    /// Set the create date/time for the document.
    ///
    /// Excel sets a date and time for every new document in UTC. The
    /// `rust_xlsxwriter` library does the same. However there may be cases
    /// where you wish to set a different creation time.  See the example above.
    ///
    /// # Parameters
    ///
    /// - `datetime`: The creation date property. A type that implements
    ///   [`IntoCustomDateTimeUtc`].
    ///
    /// [`chrono::DateTime`]:
    ///     https://docs.rs/chrono/latest/chrono/struct.DateTime.html
    ///
    pub fn set_creation_datetime(
        mut self,
        create_time: impl IntoCustomDateTimeUtc,
    ) -> DocProperties {
        self.creation_time = create_time.utc_datetime();
        self
    }

    /// Set a custom document property.
    ///
    /// Set a user defined property that will appear in the Custom section of
    /// the document properties.
    ///
    /// Excel support custom data types that are equivalent to the Rust types:
    /// [`&str`], [`f64`], [`i32`] [`bool`] and `&DateTime<Utc>`
    ///
    /// # Parameters
    ///
    /// - `name`: The user defined name of the custom property.
    /// - `value`: The value can be a [`&str`], [`f64`], [`i32`] [`bool`],
    ///   [`ExcelDateTime`] or [`chrono::DateTime<Utc>`] type for which the
    ///   `IntoCustomProperty` trait is implemented.
    ///
    /// [`chrono::DateTime<Utc>`]:
    /// https://docs.rs/chrono/latest/chrono/struct.DateTime.html
    ///
    ///
    /// # Examples
    ///
    /// An example of setting custom/user defined workbook document properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_properties_custom.rs
    /// #
    /// use rust_xlsxwriter::{DocProperties, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     let mut workbook = Workbook::new();
    ///
    ///     let properties = DocProperties::new()
    ///         .set_custom_property("Checked by", "Admin")
    ///         .set_custom_property("Cross check", true)
    ///         .set_custom_property("Department", "Finance")
    ///         .set_custom_property("Document number", 55301);
    ///
    ///     workbook.set_properties(&properties);
    ///
    ///     workbook.save("properties.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/doc_properties_custom.png">
    ///
    pub fn set_custom_property(
        mut self,
        name: impl Into<String>,
        value: impl IntoCustomProperty,
    ) -> DocProperties {
        self.custom_properties.push(value.new_custom_property(name));

        self
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------

/// The CustomProperty struct represents data types used in Excel’s custom
/// document properties.
#[doc(hidden)]
#[derive(Clone)]
pub struct CustomProperty {
    pub(crate) property_type: CustomPropertyType,
    pub(crate) name: String,
    pub(crate) text: String,
    pub(crate) number_int: i32,
    pub(crate) number_real: f64,
    pub(crate) boolean: bool,
    pub(crate) datetime: String,
}

impl Default for CustomProperty {
    fn default() -> Self {
        CustomProperty {
            property_type: CustomPropertyType::Text,
            name: String::new(),
            text: String::new(),
            number_int: 0,
            number_real: 0.0,
            boolean: true,
            datetime: ExcelDateTime::utc_now(),
        }
    }
}

impl CustomProperty {
    pub(crate) fn new_property_string(name: String, value: String) -> CustomProperty {
        CustomProperty {
            property_type: CustomPropertyType::Text,
            name,
            text: value,
            ..Default::default()
        }
    }

    pub(crate) fn new_property_i32(name: String, value: i32) -> CustomProperty {
        CustomProperty {
            property_type: CustomPropertyType::Int,
            name,
            number_int: value,
            ..Default::default()
        }
    }

    pub(crate) fn new_property_f64(name: String, value: f64) -> CustomProperty {
        CustomProperty {
            property_type: CustomPropertyType::Real,
            name,
            number_real: value,
            ..Default::default()
        }
    }

    pub(crate) fn new_property_bool(name: String, value: bool) -> CustomProperty {
        CustomProperty {
            property_type: CustomPropertyType::Bool,
            name,
            boolean: value,
            ..Default::default()
        }
    }

    pub(crate) fn new_property_datetime(
        name: String,
        value: impl IntoCustomDateTimeUtc,
    ) -> CustomProperty {
        CustomProperty {
            property_type: CustomPropertyType::DateTime,
            name,
            datetime: value.utc_datetime(),
            ..Default::default()
        }
    }
}

#[derive(Clone)]
pub(crate) enum CustomPropertyType {
    Text,
    Int,
    Real,
    Bool,
    DateTime,
}

/// Trait to map different Rust types into Excel data types used in custom document properties.
///
pub trait IntoCustomProperty {
    /// Types/objects supporting this trait must be able to convert to a
    /// [`CustomProperty`] struct.
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty;
}

impl IntoCustomProperty for &str {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_string(name.into(), self.into())
    }
}

impl IntoCustomProperty for String {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_string(name.into(), self)
    }
}

impl IntoCustomProperty for &String {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_string(name.into(), self.into())
    }
}

impl IntoCustomProperty for i32 {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_i32(name.into(), self)
    }
}

impl IntoCustomProperty for f64 {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_f64(name.into(), self)
    }
}

impl IntoCustomProperty for bool {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_bool(name.into(), self)
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoCustomProperty for &DateTime<Utc> {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_datetime(name.into(), self)
    }
}

impl IntoCustomProperty for &ExcelDateTime {
    fn new_custom_property(self, name: impl Into<String>) -> CustomProperty {
        CustomProperty::new_property_datetime(name.into(), self)
    }
}

/// Trait to map user date types to an Excel UTC date.
///
/// Map a date to the Excel UTC date used in custom document properties.
///
/// This can be either a [`ExcelDateTime`] date instance or, if the `chrono`
/// feature is enabled, a [`chrono::DateTime<Utc>`] instance.
///
/// [`chrono::DateTime<Utc>`]:
/// https://docs.rs/chrono/latest/chrono/struct.DateTime.html
///
///
pub trait IntoCustomDateTimeUtc {
    /// Trait method to convert a date into an Excel UTC date.
    ///
    fn utc_datetime(self) -> String;
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoCustomDateTimeUtc for &DateTime<Utc> {
    fn utc_datetime(self) -> String {
        self.to_rfc3339_opts(chrono::SecondsFormat::Secs, true)
    }
}

impl IntoCustomDateTimeUtc for &ExcelDateTime {
    fn utc_datetime(self) -> String {
        self.to_rfc3339()
    }
}
