// Color unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod format_tests {

    use crate::Color;

    #[test]
    fn test_hex_value() {
        assert_eq!("FF000000", Color::Default.argb_hex_value());
        assert_eq!("FF000000", Color::Black.argb_hex_value());
        assert_eq!("FF0000FF", Color::Blue.argb_hex_value());
        assert_eq!("FF800000", Color::Brown.argb_hex_value());
        assert_eq!("FF00FFFF", Color::Cyan.argb_hex_value());
        assert_eq!("FF808080", Color::Gray.argb_hex_value());
        assert_eq!("FF008000", Color::Green.argb_hex_value());
        assert_eq!("FF00FF00", Color::Lime.argb_hex_value());
        assert_eq!("FFFF00FF", Color::Magenta.argb_hex_value());
        assert_eq!("FF000080", Color::Navy.argb_hex_value());
        assert_eq!("FFFF6600", Color::Orange.argb_hex_value());
        assert_eq!("FFFFC0CB", Color::Pink.argb_hex_value());
        assert_eq!("FF800080", Color::Purple.argb_hex_value());
        assert_eq!("FFFF0000", Color::Red.argb_hex_value());
        assert_eq!("FFC0C0C0", Color::Silver.argb_hex_value());
        assert_eq!("FFFFFFFF", Color::White.argb_hex_value());
        assert_eq!("FFFFFF00", Color::Yellow.argb_hex_value());
        assert_eq!("FFABCDEF", Color::RGB(0xABCDEF).argb_hex_value());
        assert_eq!("FF000000", Color::Theme(2, 1).argb_hex_value());
    }
}
