// properties - A module for representing document properties.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use chrono::{DateTime, Utc};

/// The Properties struct is used to create an object to represent document
/// metadata properties.
///
/// It is used to create an object to represent various document properties for
/// an Excel file such as the Author's name or the Creation Date.
///
/// <img src="https://rustxlsxwriter.github.io/images/app_doc_properties.png">
///
/// The Properties struct is used in conjunction with the
/// [`workbook.set_properties()`](crate::Workbook::set_properties) method.
///
/// # Examples
///
/// An example of setting workbook document properties for a file created using
/// the rust_xlsxwriter library. This creates the file used to generate the
/// above image.
///
/// ```
/// # // This code is available in examples/app_doc_properties.rs
/// #
/// use rust_xlsxwriter::{Properties, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     let properties = Properties::new()
///         .set_title("This is an example spreadsheet")
///         .set_subject("That demonstrates document properties")
///         .set_author("A. Rust User")
///         .set_manager("J. Alfred Prufrock")
///         .set_company("Rust Solutions Inc")
///         .set_category("Sample spreadsheets")
///         .set_keywords("Sample, Example, Properties")
///         .set_comment("Created with Rust and rust_xlsxwriter");
///
///     workbook.set_properties(&properties);
///
///     let worksheet = workbook.add_worksheet();
///
///     worksheet.set_column_width(0, 30)?;
///     worksheet.write_string_only(0, 0, "See File -> Info -> Properties")?;
///
///     workbook.save("doc_properties.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// # Checksum of a saved file
///
/// A common issue that occurs with `rust_xlsxwriter`, but also with Excel, is that
/// running the same program twice doesn't generate the same file, byte for byte.
/// This can cause issue with systems that do checksumming for testing purposes.
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
///     worksheet.write_string_only(0, 0, "Hello")?;
///
///     workbook.save("properties.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// If we run this several times, with a small delay, we will get different checksums as shown below:
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
/// This is due to a file creation datetime that is included in the file and which
/// changes each time a new file is created.
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
/// use chrono::{TimeZone, Utc};
/// use rust_xlsxwriter::{Properties, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     // Create a file creation date for the file.
///     let date = Utc.with_ymd_and_hms(2023, 1, 1, 0, 0, 0).unwrap();
///
///     // Add it to the document metadata.
///     let properties = Properties::new().set_creation_datetime(&date);
///     workbook.set_properties(&properties);
///
///     let worksheet = workbook.add_worksheet();
///     worksheet.write_string_only(0, 0, "Hello")?;
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
pub struct Properties {
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
    pub(crate) creation_time: DateTime<Utc>,
}

impl Default for Properties {
    fn default() -> Self {
        Self::new()
    }
}

impl Properties {
    /// Create a new Properties struct.
    pub fn new() -> Properties {
        Properties {
            title: "".to_string(),
            status: "".to_string(),
            author: "".to_string(),
            comment: "".to_string(),
            company: "".to_string(),
            manager: "".to_string(),
            subject: "".to_string(),
            category: "".to_string(),
            keywords: "".to_string(),
            hyperlink_base: "".to_string(),
            creation_time: Utc::now(),
        }
    }

    /// Set the Title field of the document properties.
    ///
    /// Set the "Title" field of the document properties to create a title for
    /// the document such as "Sales Report".
    ///
    /// # Arguments
    ///
    /// * `title` - The title string property.
    ///
    pub fn set_title(mut self, title: &str) -> Properties {
        self.title = title.to_string();

        self
    }

    /// Set the Subject field of the document properties.
    ///
    /// Set the "Subject" field of the document properties to indicate the
    /// subject matter.
    ///
    /// # Arguments
    ///
    /// * `subject` - The subject string property.
    ///
    pub fn set_subject(mut self, subject: &str) -> Properties {
        self.subject = subject.to_string();

        self
    }

    /// Set the Manager field of the document properties.
    ///
    /// Set the "Manager" field of the document properties.
    ///
    /// # Arguments
    ///
    /// * `manager` - The manager string property.
    ///
    pub fn set_manager(mut self, manager: &str) -> Properties {
        self.manager = manager.to_string();

        self
    }

    /// Set the Company field of the document properties.
    ///
    /// Set the "Company" field of the document properties.
    ///
    /// # Arguments
    ///
    /// * `company` - The company string property.
    ///
    pub fn set_company(mut self, company: &str) -> Properties {
        self.company = company.to_string();

        self
    }

    /// Set the Category field of the document properties.
    ///
    /// Set the "Category" field of the document properties to indicate the
    /// category that the file belongs to.
    ///
    /// # Arguments
    ///
    /// * `category` - The category string property.
    ///
    pub fn set_category(mut self, category: &str) -> Properties {
        self.category = category.to_string();

        self
    }

    /// Set the Author field of the document properties.
    ///
    /// Set the "Author" field of the document properties.
    ///
    /// # Arguments
    ///
    /// * `author` - The author string property.
    ///
    pub fn set_author(mut self, author: &str) -> Properties {
        self.author = author.to_string();

        self
    }

    /// Set the Keywords field of the document properties.
    ///
    /// Set the "Keywords" field of the document properties. This can be one or
    /// more keywords that can be used in searches.
    ///
    /// # Arguments
    ///
    /// * `keywords` - The keywords string property.
    ///
    pub fn set_keywords(mut self, keywords: &str) -> Properties {
        self.keywords = keywords.to_string();

        self
    }

    /// Set the Comment field of the document properties.
    ///
    /// Set the "Comment" field of the document properties. This can be a
    /// general comment or summary that you want to add to the properties.
    ///
    /// # Arguments
    ///
    /// * `comment` - The comment string property.
    ///
    pub fn set_comment(mut self, comment: &str) -> Properties {
        self.comment = comment.to_string();

        self
    }

    /// Set the Status field of the document properties.
    ///
    /// Set the "Status" field of the document properties such as "Draft" or
    /// "Final".
    ///
    /// # Arguments
    ///
    /// * `status` - The status string property.
    ///
    pub fn set_status(mut self, status: &str) -> Properties {
        self.status = status.to_string();

        self
    }

    /// Set the Hyperlink_base field of the document properties.
    ///
    /// Set the "Hyperlink_base" field of the document properties to have a
    /// default base url.
    ///
    /// # Arguments
    ///
    /// * `hyperlink_base` - The hyperlink base string property.
    ///
    pub fn set_hyperlink_base(mut self, hyperlink_base: &str) -> Properties {
        self.hyperlink_base = hyperlink_base.to_string();

        self
    }

    /// Set the create date/time for the document.
    ///
    /// Excel sets a date and time for every new document in UTC. The
    /// `rust_xlsxwriter` library does the same. However there may be cases
    /// where you wish to set a different creation time.
    ///
    /// # Arguments
    ///
    /// * `datetime` - The hyperlink_base string property. [`chrono::DateTime`]
    ///
    /// [`chrono::DateTime`]:
    ///     https://docs.rs/chrono/latest/chrono/struct.DateTime.html
    ///
    pub fn set_creation_datetime(mut self, create_time: &DateTime<Utc>) -> Properties {
        self.creation_time = *create_time;

        self
    }
}
