// Format unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod format_tests {

    use crate::{
        Color, FontScheme, Format, FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern,
        FormatUnderline,
    };
    use pretty_assertions::assert_eq;

    #[test]
    fn test_unset() {
        let format1 = Format::default();
        let format2 = Format::new()
            .set_bold()
            .set_italic()
            .set_font_strikethrough()
            .set_text_wrap()
            .set_shrink()
            .set_unlocked()
            .set_hidden()
            .set_quote_prefix()
            .set_checkbox()
            // Unset the properties.
            .unset_bold()
            .unset_italic()
            .unset_font_strikethrough()
            .unset_text_wrap()
            .unset_shrink()
            .set_locked()
            .unset_hidden()
            .unset_quote_prefix()
            .unset_checkbox();

        assert_eq!(format1, format2);
    }

    #[test]
    fn test_merge_num_format() {
        let default = Format::new();
        let has_value = Format::new().set_num_format("0.000");

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }

    #[test]
    fn test_merge_num_format_index() {
        let default = Format::new();
        let has_value = Format::new().set_num_format_index(1);

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }

    #[test]
    fn test_merge_fill() {
        let default = Format::new();
        let has_value = Format::new()
            .set_background_color(Color::Yellow)
            .set_foreground_color(Color::Orange)
            .set_pattern(FormatPattern::Gray0625);

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }

    #[test]
    fn test_merge_font() {
        let default = Format::new();
        let has_value = Format::new()
            .set_bold()
            .set_italic()
            .set_underline(FormatUnderline::DoubleAccounting)
            .set_font_name("Name")
            .set_font_size(42)
            .set_font_charset(7)
            .set_font_family(6)
            .set_font_script(crate::FormatScript::Subscript)
            .set_font_strikethrough()
            .set_font_color(Color::Red)
            .set_font_scheme(FontScheme::Headings);

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }

    #[test]
    fn test_merge_border() {
        let default = Format::new();
        let has_value = Format::new()
            .set_border_top_color(Color::Red)
            .set_border_left_color(Color::Green)
            .set_border_right_color(Color::White)
            .set_border_bottom_color(Color::Yellow)
            .set_border_top(FormatBorder::DashDotDot)
            .set_border_left(FormatBorder::Dotted)
            .set_border_right(FormatBorder::DashDot)
            .set_border_bottom(FormatBorder::Medium)
            .set_border_diagonal_color(Color::Black)
            .set_border_diagonal(FormatBorder::SlantDashDot)
            .set_border_diagonal_type(FormatDiagonalBorder::BorderUpDown);

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }

    #[test]
    fn test_merge_alignment() {
        let default = Format::new();
        let has_value = Format::new()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_shrink()
            .set_text_wrap()
            .set_rotation(45)
            .set_indent(1)
            .set_reading_direction(2);

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }

    #[test]
    fn test_merge_misc() {
        let default = Format::new();
        let has_value = Format::new()
            .set_hidden()
            .set_unlocked()
            .set_quote_prefix()
            .set_checkbox();

        // Test an overwriting merge.
        let merged = default.merge(&has_value);
        assert_eq!(merged, has_value);

        // Test a non-overwriting merge.
        let merged = has_value.merge(&default);
        assert_eq!(merged, has_value);
    }
}
