// url - A module for representing Excel worksheet URLs.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::{XlsxError, MAX_PARAMETER_LEN};

const MAX_URL_LEN: usize = 2_080;

/// The `Url` struct is used to define a worksheet URL.
///
/// The `Url` struct creates a URL type that can be used to write worksheet
/// URLs.
///
/// In general, you would use the
/// [`Worksheet::write_url()`](crate::Worksheet::write_url) with a string
/// representation of the URL, like this:
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
/// #     // Write a URL with a simple string argument.
///     worksheet.write_url(0, 0, "https://www.rust-lang.org")?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// The URL will then be displayed as expected in Excel:
///
/// <img src="https://rustxlsxwriter.github.io/images/url_intro1.png">
///
/// To differentiate a URL from an ordinary string (for example, when
/// storing it in a data structure), you can also represent the URL with a
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
/// #     // Write a URL with a Url struct.
///     worksheet.write_url(0, 0, Url::new("https://www.rust-lang.org"))?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Using a `Url` struct also allows you to write a URL using the generic
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
/// #     // Write a URL with a Url struct and generic write().
///     worksheet.write(0, 0, Url::new("https://www.rust-lang.org"))?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// There are three types of URLs/links supported by Excel and the `rust_xlsxwriter`
/// library:
///
/// 1. Web-based URIs like:
///
///    * `http://`, `https://`, `ftp://`, `ftps://`, and `mailto:`.
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
///    other Excel files, the URL string can include a sheet and cell reference
///    after the `"#"` anchor, as shown in the last two examples above. When using
///    Windows paths, like in the examples above, it is best to use a Rust raw
///    string to avoid issues with the backslashes:
///    `r"file:///C:\Temp\Book1.xlsx"`.
///
/// 3. Internal links to a cell or range of cells in the workbook using the
///    pseudo-URI `internal:`:
///
///    * `internal:Sheet2!A1`
///    * `internal:Sheet2!A1:G5`
///    * `internal:'Sales Data'!A1`
///
///    Worksheet references are typically of the form `Sheet1!A1`, where a
///    worksheet and target cell should be specified. You can also link to a
///    worksheet range using the standard Excel range notation like
///    `Sheet1!A1:B2`. Excel requires that worksheet names containing spaces or
///    non-alphanumeric characters are single-quoted as follows: `'Sales
///    Data'!A1`.
///
/// The library will escape the following characters in URLs as required by
/// Excel: ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
/// style escapes. In that case, it is assumed that the URL was escaped
/// correctly by the user and will be passed directly to Excel.
///
/// Excel has a limit of around 2080 characters in the URL string. URLs beyond
/// this limit will raise an error when written.
///
/// # Examples
///
/// An example of some of the features of URLs/hyperlinks.
///
/// ```
/// # // This code is available in examples/app_hyperlinks.rs
/// #
/// use rust_xlsxwriter::{Color, Format, FormatUnderline, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Create a format to use in the worksheet.
///     let link_format = Format::new()
///         .set_font_color(Color::Red)
///         .set_underline(FormatUnderline::Single);
///
///     // Add a worksheet to the workbook.
///     let worksheet1 = workbook.add_worksheet();
///
///     // Set the column width for clarity.
///     worksheet1.set_column_width(0, 26)?;
///
///     // Write some URL links.
///     worksheet1.write_url(0, 0, "https://www.rust-lang.org")?;
///     worksheet1.write_url_with_text(1, 0, "https://www.rust-lang.org", "Learn Rust")?;
///     worksheet1.write_url_with_format(2, 0, "https://www.rust-lang.org", &link_format)?;
///
///     // Write some internal links.
///     worksheet1.write_url(4, 0, "internal:Sheet1!A1")?;
///     worksheet1.write_url(5, 0, "internal:Sheet2!C4")?;
///
///     // Write some external links.
///     worksheet1.write_url(7, 0, r"file:///C:\Temp\Book1.xlsx")?;
///     worksheet1.write_url(8, 0, r"file:///C:\Temp\Book1.xlsx#Sheet1!C4")?;
///
///     // Add another sheet to link to.
///     let worksheet2 = workbook.add_worksheet();
///     worksheet2.write_string(3, 2, "Here I am")?;
///     worksheet2.write_url_with_text(4, 2, "internal:Sheet1!A6", "Go back")?;
///
///     // Save the file to disk.
///     workbook.save("hyperlinks.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/app_hyperlinks.png">
///
/// This is an example of creating a "Table of Contents" worksheet with links to
/// other worksheets in the workbook.
///
/// ```
/// # // This code is available in examples/app_table_of_contents.rs
/// #
/// use rust_xlsxwriter::{utility::quote_sheet_name, Format, Url, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Create a table of contents worksheet at the start. If the worksheet names
///     // are known in advance you can do add them here. For the sake of this
///     // example we will assume that they aren't known and/or are created
///     // dynamically.
///     let _ = workbook.add_worksheet().set_name("Overview")?;
///
///     // Add some worksheets.
///     let _ = workbook.add_worksheet().set_name("Pricing")?;
///     let _ = workbook.add_worksheet().set_name("Sales")?;
///     let _ = workbook.add_worksheet().set_name("Revenue")?;
///     let _ = workbook.add_worksheet().set_name("Analytics")?;
///
///     // If the sheet names aren't known in advance we can find them as follows:
///     let mut worksheet_names = workbook
///         .worksheets()
///         .iter()
///         .map(|worksheet| worksheet.name())
///         .collect::<Vec<_>>();
///
///     // Remove the "Overview" worksheet name.
///     worksheet_names.remove(0);
///
///     // Get the "Overview" worksheet to add the table of contents.
///     let worksheet = workbook.worksheet_from_name("Overview")?;
///
///     // Write a header.
///     let header = Format::new().set_bold().set_background_color("C6EFCE");
///     worksheet.write_string_with_format(0, 0, "Table of Contents", &header)?;
///
///     // Write the worksheet names with links.
///     for (i, name) in worksheet_names.iter().enumerate() {
///         let sheet_name = quote_sheet_name(name);
///         let link = format!("internal:{sheet_name}!A1");
///         let url = Url::new(link).set_text(name);
///
///         worksheet.write_url(i as u32 + 1, 0, &url)?;
///     }
///
///     // Autofit the data for clarity.
///     worksheet.autofit();
///
///     // Save the file to disk.
///     workbook.save("table_of_contents.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/app_table_of_contents.png">
///
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
            url_link: link.clone(),            // The worksheet hyperlink URL.
            user_text: String::new(),          // Text the user sees. May be the same as the URL.
            rel_link: link.clone(),            // The URL as it appears in a relationship file.
            rel_anchor: String::new(),         // Equivalent to a URL anchor_fragment.
            rel_display: false,                // Relationship display setting.
            rel_id: 0,                         // Relationship id.
            tool_tip: String::new(),           // The mouseover tool tip.
            link_type: HyperlinkType::Unknown, // Url, file, internal.
        }
    }

    /// Set the alternative text for the URL.
    ///
    /// Set an alternative, user friendly, text for the URL.
    ///
    /// # Parameters
    ///
    /// `text` - The alternative text, as a string or string like type.
    ///
    /// # Examples
    ///
    /// The following example demonstrates writing a URL to a worksheet with
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
    /// #     // Write a URL with a Url struct and alternative text.
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

    /// Set the screen tip for the URL.
    ///
    /// Set a screen tip when the user does a mouseover of the URL.
    ///
    /// # Parameters
    ///
    /// `tip` - The URL tip, as a string or string like type.
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

        // Check the URL string lengths are within Excel's limits. The user text
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

            // Split the URL into URL + #anchor if that exists.
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

            // Split the URL into URL + #anchor if that exists.
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
        // Escape any URL characters in the URL string.
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
