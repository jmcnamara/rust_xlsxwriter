// hyperlink - A struct to represent Excel hyperlinks.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::{static_regex, Url, XlsxError, MAX_PARAMETER_LEN};

const MAX_URL_LEN: usize = 2_080;

/// A struct to represent different Excel hyperlinks types.
///
/// Hyperlinks are used in worksheets cells and also in images (and although not
/// currently supported) other shapes and objects. In general hyperlinks are
/// stored as a `rel_id` (relationship id) that references a `_rels/*.xml.rels`
/// file where the actual link is stored.
///
/// There are 3 types of url/link supported by Excel:
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
/// The struct will escape the following characters in URLs as required by
/// Excel, ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
/// style escapes. In which case it is assumed that the URL was escaped
/// correctly by the user and will by passed directly to Excel.
///
/// Excel has a limit of around 2080 characters in the url string. Urls beyond
/// this limit will raise an error when written.
///
#[derive(Clone, Debug)]
pub(crate) struct Hyperlink {
    pub(crate) url_link: String,
    pub(crate) rel_link: String,
    pub(crate) user_text: String,
    pub(crate) tool_tip: String,
    pub(crate) rel_anchor: String,
    pub(crate) rel_display: bool,
    pub(crate) link_type: HyperlinkType,
    pub(crate) rel_id: u32,
}

impl Hyperlink {
    pub(crate) fn new(url: &Url) -> Result<Hyperlink, XlsxError> {
        let mut hyperlink = Hyperlink {
            url_link: url.link.clone(),        // The worksheet hyperlink url.
            user_text: url.text.clone(),       // Text the user sees. May be the same as the url.
            rel_link: url.link.clone(),        // The url as it appears in a relationship file.
            rel_anchor: String::new(),         // Equivalent to a url anchor_fragment.
            rel_display: false,                // Relationship display setting.
            rel_id: 0,                         // Relationship id.
            tool_tip: url.tip.clone(),         // The mouseover tool tip.
            link_type: HyperlinkType::Unknown, // Url, file, internal.
        };

        Self::parse_url(&mut hyperlink)?;

        // Check the hyperlink string lengths are within Excel's limits. The
        // user text length is checked by write_string_with_format().
        if hyperlink.url_link.chars().count() > MAX_URL_LEN
            || hyperlink.rel_anchor.chars().count() > MAX_URL_LEN
        {
            return Err(XlsxError::MaxUrlLengthExceeded);
        }

        // Escape hyperlink strings after length checks.
        Self::escape_strings(&mut hyperlink);

        if hyperlink.tool_tip.chars().count() > MAX_PARAMETER_LEN {
            return Err(XlsxError::ParameterError(
                "Hyperlink tool tip must be less than or equal to Excel's limit of characters"
                    .to_string(),
            ));
        }

        Ok(hyperlink)
    }

    // This method handles a variety of different string processing required for
    // links and targets associated with Excel's hyperlinks.
    fn parse_url(&mut self) -> Result<(), XlsxError> {
        let remote_file = static_regex!(r"^(\\\\|\w:)");
        let url_protocol = static_regex!(r"^(ftp|http)s?://");
        let original_url = self.url_link.clone();

        if url_protocol.is_match(&self.url_link) {
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

            // Links to local files aren't prefixed with file:///.
            if !remote_file.is_match(&link_path) {
                self.url_link.clone_from(&link_path);
            }

            self.rel_link.clone_from(&self.url_link);
            if !remote_file.is_match(&link_path) {
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
    fn escape_strings(&mut self) {
        let url_escape = static_regex!(r"%[0-9a-fA-F]{2}");

        // Escape any url characters in the url string.
        if !url_escape.is_match(&self.url_link) {
            self.url_link = crate::xmlwriter::escape_url(&self.url_link).into();
        }

        // Escape the link used in the relationship file, except for internal
        // links which are generally sheet/cell locations.
        if self.link_type != HyperlinkType::Internal && !url_escape.is_match(&self.rel_link) {
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
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum HyperlinkType {
    Unknown,
    Url,
    Internal,
    File,
}
