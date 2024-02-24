// url - A module for representing Excel worksheet Urls.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

/// The `Url` struct is used to define a worksheet url.
///
/// The `Url` struct creates a url type that can be used to write worksheet
/// urls.
///
/// In general you would use the
/// [`worksheet.write_url()`](crate::Worksheet::write_url) with a string
/// representation of the url, like this:
///
/// ```
/// # // This code is available in examples/doc_url_intro1.rs
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
/// #     // Write a url with a simple string argument.
///     worksheet.write_url(0, 0, "https://www.rust-lang.org")?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// The url will then be displayed as expected in Excel:
///
/// <img src="https://rustxlsxwriter.github.io/images/url_intro1.png">
///
/// In order to differentiate a url from an ordinary string (for example when
/// storing it in a data structure) you can also represent the url with a
/// [`Url`] struct:
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
/// #     // Write a url with a Url struct.
///     worksheet.write_url(0, 0, Url::new("https://www.rust-lang.org"))?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Using a `Url` struct also allows you to write a url using the generic
/// [`worksheet.write()`](crate::Worksheet::write) method:
///
/// ```
/// # // This code is available in examples/doc_url_intro3.rs
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
/// #     // Write a url with a Url struct and generic write().
///     worksheet.write(0, 0, Url::new("https://www.rust-lang.org"))?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// There are 3 types of url/link supported by Excel and the `rust_xlsxwriter`
/// library:
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
///    other Excel files the url string can include a sheet and cell reference
///    after the `"#"` anchor, as shown in the last 2 examples above. When using
///    Windows paths, like in the examples above, it is best to use a Rust raw
///    string to avoid issues with the backslashes:
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
///    `Sheet1!A1:B2`. Excel requires that worksheet names containing spaces or
///    non alphanumeric characters are single quoted as follows `'Sales
///    Data'!A1`.
///
/// The library will escape the following characters in URLs as required by
/// Excel, ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
/// style escapes. In which case it is assumed that the URL was escaped
/// correctly by the user and will by passed directly to Excel.
///
/// Excel has a limit of around 2080 characters in the url string. Urls beyond
/// this limit will raise an error when written.
///
#[derive(Clone, Debug)]
pub struct Url {
    pub(crate) link: String,
    pub(crate) text: String,
    pub(crate) tip: String,
}

impl Url {
    /// Create a new Url struct.
    ///
    /// # Parameters
    ///
    /// `link` - A string like type representing a URL.
    ///
    pub fn new(link: impl Into<String>) -> Url {
        Url {
            link: link.into(),
            text: String::new(),
            tip: String::new(),
        }
    }

    /// Set the alternative text for the url.
    ///
    /// Set an alternative, user friendly, text for the url.
    ///
    /// # Parameters
    ///
    /// `text` - The alternative text, as a string or string like type.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a url to a worksheet with
    /// alternative text.
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
    /// #     // Write a url with a Url struct and alternative text.
    ///     worksheet.write(0, 0, Url::new("https://www.rust-lang.org").set_text("Learn Rust"))?;
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
    /// <img src="https://rustxlsxwriter.github.io/images/url_set_text.png">
    ///
    pub fn set_text(mut self, text: impl Into<String>) -> Url {
        self.text = text.into();
        self
    }

    /// Set the screen tip for the url.
    ///
    /// Set a screen tip when the user does a mouseover of the url.
    ///
    /// # Parameters
    ///
    /// `tip` - The url tip, as a string or string like type.
    ///
    pub fn set_tip(mut self, tip: impl Into<String>) -> Url {
        self.tip = tip.into();
        self
    }
}

impl From<&str> for Url {
    fn from(value: &str) -> Url {
        Url::new(value)
    }
}
