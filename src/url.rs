// url - A module for representing Excel worksheet Urls.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::{XlsxError, MAX_PARAMETER_LEN};

const MAX_URL_LEN: usize = 2_080;

/// The `Url` struct is used to define a worksheet url.
///
/// The `Url` struct creates a url type that can be used to write worksheet
/// urls.
///
/// In general you would use the
/// [`Worksheet::write_url()`](crate::Worksheet::write_url) with a string
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
/// [`Worksheet::write()`](crate::Worksheet::write) method:
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
    pub(crate) url_link: String,
    pub(crate) rel_link: String,
    pub(crate) user_text: String,
    pub(crate) tool_tip: String,
    pub(crate) rel_anchor: String,
    pub(crate) rel_display: bool,
    pub(crate) link_type: HyperlinkType,
    pub(crate) rel_id: u32,
}

impl Url {
    /// Create a new Url struct.
    ///
    /// # Parameters
    ///
    /// `link` - A string like type representing a URL.
    ///
    pub fn new(link: impl Into<String>) -> Url {
        let link = link.into();

        Url {
            url_link: link.clone(),            // The worksheet hyperlink url.
            user_text: String::new(),          // Text the user sees. May be the same as the url.
            rel_link: link.clone(),            // The url as it appears in a relationship file.
            rel_anchor: String::new(),         // Equivalent to a url anchor_fragment.
            rel_display: false,                // Relationship display setting.
            rel_id: 0,                         // Relationship id.
            tool_tip: String::new(),           // The mouseover tool tip.
            link_type: HyperlinkType::Unknown, // Url, file, internal.
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
        self.user_text = text.into();
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
        self.tool_tip = tip.into();
        self
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    pub(crate) fn initialize(&mut self) -> Result<(), XlsxError> {
        self.parse_url()?;

        // Check the url string lengths are within Excel's limits. The user text
        // length is checked by write_string_with_format().
        if self.url_link.chars().count() > MAX_URL_LEN
            || self.rel_anchor.chars().count() > MAX_URL_LEN
        {
            return Err(XlsxError::MaxUrlLengthExceeded);
        }

        // Escape hyperlink strings after length checks.
        self.escape_strings();

        if self.tool_tip.chars().count() > MAX_PARAMETER_LEN {
            return Err(XlsxError::ParameterError(
                "Hyperlink tool tip must be less than or equal to Excel's limit of characters"
                    .to_string(),
            ));
        }

        Ok(())
    }

    // This method handles a variety of different string processing required for
    // links and targets associated with Excel's urls/hyperlinks.
    pub(crate) fn parse_url(&mut self) -> Result<(), XlsxError> {
        let original_url = self.url_link.clone();

        if self.url_link.starts_with("http://")
            || self.url_link.starts_with("https://")
            || self.url_link.starts_with("ftp://")
            || self.url_link.starts_with("ftps://")
        {
            // Handle web links like https://.
            self.link_type = HyperlinkType::Url;

            if self.user_text.is_empty() {
                self.user_text.clone_from(&self.url_link);
            }

            // Split the url into url + #anchor if that exists.
            let parts: Vec<&str> = self.url_link.splitn(2, '#').collect();
            if parts.len() == 2 {
                self.rel_anchor = parts[1].to_string();
                self.url_link = parts[0].to_string();
            }
        } else if self.url_link.starts_with("mailto:") {
            // Handle mail address links.
            self.link_type = HyperlinkType::Url;

            if self.user_text.is_empty() {
                self.user_text = self.url_link.replacen("mailto:", "", 1);
            }
        } else if self.url_link.starts_with("internal:") {
            // Handle links to cells within the workbook.
            self.link_type = HyperlinkType::Internal;
            self.rel_anchor = self.url_link.replacen("internal:", "", 1);

            if self.user_text.is_empty() {
                self.user_text.clone_from(&self.rel_anchor);
            }
        } else if self.url_link.starts_with("file://") {
            // Handle links to other files or cells in other Excel files.
            self.link_type = HyperlinkType::File;
            let link_path = self.url_link.replacen("file:///", "", 1);
            let link_path = link_path.replacen("file://", "", 1);

            // Links to relative file paths should be stored without file:///.
            let is_relative_path = Self::relative_path(&link_path);
            if is_relative_path {
                self.url_link.clone_from(&link_path);
            }

            // Links to relative file paths should continue to use Windows "\"
            // path separator. Other paths should use "/".
            self.rel_link.clone_from(&self.url_link);
            if is_relative_path {
                self.rel_link = self.rel_link.replace('\\', "/");
            }

            if self.user_text.is_empty() {
                self.user_text = link_path;
            }

            // Split the url into url + #anchor if that exists.
            let parts: Vec<&str> = self.url_link.splitn(2, '#').collect();
            if parts.len() == 2 {
                self.rel_anchor = parts[1].to_string();
                self.url_link = parts[0].to_string();
            }
        } else {
            return Err(XlsxError::UnknownUrlType(original_url));
        }

        Ok(())
    }

    // Escape hyperlink string variants.
    pub(crate) fn escape_strings(&mut self) {
        // Escape any url characters in the url string.
        if !Self::is_escaped(&self.url_link) {
            self.url_link = crate::xmlwriter::escape_url(&self.url_link).into();
        }

        // Escape the link used in the relationship file, except for internal
        // links which are generally sheet/cell locations.
        if self.link_type != HyperlinkType::Internal && !Self::is_escaped(&self.rel_link) {
            self.rel_link = crate::xmlwriter::escape_url(&self.rel_link).into();
        }

        // Excel additionally escapes # to %23 in file paths.
        if self.link_type == HyperlinkType::File {
            self.rel_link = self.rel_link.replace('#', "%23");
        }
    }

    // Increment the relationship id for some types only.
    pub(crate) fn increment_rel_id(&mut self, rel_id: u32) -> u32 {
        match self.link_type {
            HyperlinkType::Url | HyperlinkType::File => {
                self.rel_id = rel_id;
                rel_id + 1
            }
            _ => rel_id,
        }
    }

    // Get the target for relationship ids.
    pub(crate) fn target(&mut self) -> String {
        let mut target = self.rel_link.clone();

        if self.link_type == HyperlinkType::Internal {
            target = target.replace("internal:", "#");
        }

        target
    }

    // Get the target mode for relationship ids.
    pub(crate) fn target_mode(&mut self) -> String {
        let mut target_mode = String::from("External");

        if self.link_type == HyperlinkType::Internal {
            target_mode = String::new();
        }

        target_mode
    }

    // Check for relative paths like "file.xlsx" or "../file.xlsx" as anything that
    // isn't an absolute path like "\\share\file.xlsx" or "C:\temp\file.xlsx".
    pub(crate) fn relative_path(url: &str) -> bool {
        // Check for Windows network share links.
        if url.starts_with(r"\\") {
            return false;
        }

        // Check for Windows path links like C:\temp\file.xlsx.
        if let Some(position) = url.find(':') {
            if position == 1 && url.starts_with(|c: char| c.is_ascii()) {
                return false;
            }
        }

        true
    }

    // Check if a URL string is already HTML escaped.
    pub(crate) fn is_escaped(url: &str) -> bool {
        if !url.contains('%') {
            return false;
        }

        url.as_bytes().windows(3).any(|w| {
            matches!(
                w,
                b"%25"
                    | b"%22"
                    | b"%20"
                    | b"%3c"
                    | b"%3e"
                    | b"%5b"
                    | b"%5d"
                    | b"%5e"
                    | b"%60"
                    | b"%7b"
                    | b"%7d"
            )
        })
    }
}

// -----------------------------------------------------------------------
// Traits.
// -----------------------------------------------------------------------

impl From<&str> for Url {
    fn from(value: &str) -> Url {
        Url::new(value)
    }
}

impl From<&Url> for Url {
    fn from(value: &Url) -> Url {
        value.clone()
    }
}

// -----------------------------------------------------------------------
// HyperlinkType enum.
// -----------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum HyperlinkType {
    Unknown,
    Url,
    Internal,
    File,
}
