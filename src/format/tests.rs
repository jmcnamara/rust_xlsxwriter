// Format unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod format_tests {

    use crate::Format;

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
            .unset_bold()
            .unset_italic()
            .unset_font_strikethrough()
            .unset_text_wrap()
            .unset_shrink()
            .set_locked()
            .unset_hidden();

        assert_eq!(format1, format2);
    }
}
