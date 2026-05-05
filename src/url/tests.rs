// Url unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod url_tests {

    use crate::Url;

    #[test]
    fn test_is_escaped_valid_sequences() {
        // Test all valid escape sequences that should return true.
        let valid_escapes = [
            "%20", // Space
            "%22", // Double quote "
            "%25", // Percent sign %
            "%60", // Backtick `
            "%3c", "%3C", // Less than <
            "%3e", "%3E", // Greater than >
            "%5b", "%5B", // Left bracket [
            "%5d", "%5D", // Right bracket ]
            "%5e", "%5E", // Caret ^
            "%7b", "%7B", // Left brace {
            "%7d", "%7D", // Right brace }
        ];

        for escape in &valid_escapes {
            // Test escape sequence alone.
            assert!(Url::is_escaped(escape), "Failed for: {}", escape);

            // Test escape sequence in a URL.
            let url = format!("http://example.com/test{}", escape);
            assert!(Url::is_escaped(&url), "Failed for URL: {}", url);
        }
    }

    #[test]
    fn test_is_escaped_no_percent() {
        // URLs without percent signs should return false.
        assert!(!Url::is_escaped("https://www.example.com"));
        assert!(!Url::is_escaped("http://example.com/path"));
        assert!(!Url::is_escaped("file:///C:/temp/file.xlsx"));
        assert!(!Url::is_escaped("internal:Sheet1!A1"));
        assert!(!Url::is_escaped(""));
        assert!(!Url::is_escaped("simple string"));
    }

    #[test]
    fn test_is_escaped_percent_not_escape_sequence() {
        // URLs with percent but not valid escape sequences should return false.
        assert!(!Url::is_escaped("http://example.com/test%"));
        assert!(!Url::is_escaped("http://example.com/test%1"));
        assert!(!Url::is_escaped("http://example.com/test%AB"));
        assert!(!Url::is_escaped("http://example.com/test%99"));
        assert!(!Url::is_escaped("http://example.com/test%FF"));
        assert!(!Url::is_escaped("discount%off"));
    }
}
