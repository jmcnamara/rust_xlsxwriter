// color - A module for representing Excel color types.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

/// The `Color` enum defines Excel colors that can be used throughout the
/// `rust_xlsxwriter` APIs.
///
/// There are 3 types of colors within the enum:
///
/// 1. Predefined named colors like `Color::Green`.
/// 2. User defined RGB colors such as `Color::RGB(0x4F026A)` using a format
///    similar to html colors like `#RRGGBB`, except as an integer.
/// 3. Theme colors from the standard palette of 60 colors like `Color::Theme(9,
///    4)`. The theme colors are shown in the image below.
///
///    <img
///    src="https://rustxlsxwriter.github.io/images/theme_color_palette.png">
///
///    The syntax for theme colors in `Color` is `Theme(color, shade)` where
///    `color` is one of the 0-9 values on the top row and `shade` is the
///    variant in the associated column from 0-5. For example "White, background
///    1" in the top left is `Theme(0, 0)` and "Orange, Accent 6, Darker 50%" in
///    the bottom right is `Theme(9, 5)`.
///
/// Note, there are no plans to support anything other than the default Excel
/// "Office" theme.
///
/// # Examples
///
/// The following example demonstrates using different `Color` enum values to
/// set the color of some text in a worksheet.
///
/// ```
/// # // This code is available in examples/doc_enum_xlsxcolor.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.set_column_width(0, 14)?;
///
///     let format1 = Format::new().set_font_color(Color::Red);
///     let format2 = Format::new().set_font_color(Color::Green);
///     let format3 = Format::new().set_font_color(Color::RGB(0x4F026A));
///     let format4 = Format::new().set_font_color(Color::RGB(0x73CC5F));
///     let format5 = Format::new().set_font_color(Color::Theme(4, 0));
///     let format6 = Format::new().set_font_color(Color::Theme(9, 4));
///
///     worksheet.write_string_with_format(0, 0, "Red", &format1)?;
///     worksheet.write_string_with_format(1, 0, "Green", &format2)?;
///     worksheet.write_string_with_format(2, 0, "#4F026A", &format3)?;
///     worksheet.write_string_with_format(3, 0, "#73CC5F", &format4)?;
///     worksheet.write_string_with_format(4, 0, "Theme (4, 0)", &format5)?;
///     worksheet.write_string_with_format(5, 0, "Theme (9, 4)", &format6)?;
///
/// #     workbook.save("colors.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/enum_xlsxcolor.png">
///
/// An example of the different types of color syntax that is supported by the
/// APIs that accept [`Color`] values and that support the `impl Into<Color>`
/// trait.
///
/// ```
/// # // This code is available in examples/doc_into_color.rs
/// #
/// use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet.
///     let worksheet = workbook.add_worksheet();
///
///     // Widen the column for clarity.
///     worksheet.set_column_width_pixels(0, 80)?;
///
///     // Some examples with named color enum values.
///     let color_format = Format::new().set_background_color(Color::Green);
///     worksheet.write_string(0, 0, "Green")?;
///     worksheet.write_blank(0, 1, &color_format)?;
///
///     let color_format = Format::new().set_background_color(Color::Red);
///     worksheet.write_string(1, 0, "Red")?;
///     worksheet.write_blank(1, 1, &color_format)?;
///
///     // Write a RGB color using the Color::RGB() enum method.
///     let color_format = Format::new().set_background_color(Color::RGB(0xFF7F50));
///     worksheet.write_string(2, 0, "#FF7F50")?;
///     worksheet.write_blank(2, 1, &color_format)?;
///
///     // Write a RGB color with the shorter Html string variant.
///     let color_format = Format::new().set_background_color("#6495ED");
///     worksheet.write_string(3, 0, "#6495ED")?;
///     worksheet.write_blank(3, 1, &color_format)?;
///
///     // Write a RGB color with a Html string (but without the `#`).
///     let color_format = Format::new().set_background_color("DCDCDC");
///     worksheet.write_string(4, 0, "#DCDCDC")?;
///     worksheet.write_blank(4, 1, &color_format)?;
///
///     // Write a RGB color with the optional u32 variant.
///     let color_format = Format::new().set_background_color(0xDAA520);
///     worksheet.write_string(5, 0, "#DAA520")?;
///     worksheet.write_blank(5, 1, &color_format)?;
///
///     // Add a Theme color.
///     let color_format = Format::new().set_background_color(Color::Theme(4, 3));
///     worksheet.write_string(6, 0, "Theme(4, 3)")?;
///     worksheet.write_blank(6, 1, &color_format)?;
///
///     // Save the file to disk.
///     workbook.save("into_color.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/into_color.png">
///
#[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, Default)]
pub enum Color {
    /// A user defined RGB color in the range 0x000000 (black) to 0xFFFFFF
    /// (white). Any values outside this range will be ignored with a a warning.
    RGB(u32),

    /// A theme color on the default palette (see the image above). The syntax
    /// for theme colors is `Theme(color, shade)` where `color` is one of the
    /// 0-9 values on the top row and `shade` is the variant in the associated
    /// column from 0-5. Any values outside these ranges will be ignored with a
    /// a warning.
    Theme(u8, u8),

    /// The default color for an Excel property.
    #[default]
    Default,

    /// The Automatic color for an Excel property. This is usually the same as
    /// the `Default` color but can vary according to system settings.
    Automatic,

    /// The color Black with a RGB value of 0x000000.
    Black,

    /// The color Blue with a RGB value of 0x0000FF.
    Blue,

    /// The color Brown with a RGB value of 0x800000.
    Brown,

    /// The color Cyan with a RGB value of 0x00FFFF.
    Cyan,

    /// The color Gray with a RGB value of 0x808080.
    Gray,

    /// The color Green with a RGB value of 0x008000.
    Green,

    /// The color Lime with a RGB value of 0x00FF00.
    Lime,

    /// The color Magenta with a RGB value of 0xFF00FF.
    Magenta,

    /// The color Navy with a RGB value of 0x000080.
    Navy,

    /// The color Orange with a RGB value of 0xFF6600.
    Orange,

    /// The color Pink with a RGB value of 0xFFC0CB.
    Pink,

    /// The color Purple with a RGB value of 0x800080.
    Purple,

    /// The color Red with a RGB value of 0xFF0000.
    Red,

    /// The color Silver with a RGB value of 0xC0C0C0.
    Silver,

    /// The color White with a RGB value of 0xFFFFFF.
    White,

    /// The color Yellow with a RGB value of 0xFFFF00
    Yellow,
}

impl Color {
    // Get the RGB hex value for a color.
    pub(crate) fn rgb_hex_value(self) -> String {
        match self {
            Color::Red => "FF0000".to_string(),
            Color::Blue => "0000FF".to_string(),
            Color::Cyan => "00FFFF".to_string(),
            Color::Gray => "808080".to_string(),
            Color::Lime => "00FF00".to_string(),
            Color::Navy => "000080".to_string(),
            Color::Pink => "FFC0CB".to_string(),
            Color::Brown => "800000".to_string(),
            Color::Green => "008000".to_string(),
            Color::White => "FFFFFF".to_string(),
            Color::Orange => "FF6600".to_string(),
            Color::Purple => "800080".to_string(),
            Color::Silver => "C0C0C0".to_string(),
            Color::Yellow => "FFFF00".to_string(),
            Color::Magenta => "FF00FF".to_string(),
            Color::RGB(color) => format!("{color:06X}"),

            // Default to black for non RGB colors.
            Color::Theme(_, _) | Color::Default | Color::Automatic | Color::Black => {
                "000000".to_string()
            }
        }
    }

    // Get the RGB hex value for a VML fill color in "#rrggbb" format.
    pub(crate) fn vml_rgb_hex_value(self) -> String {
        match self {
            // Use Comment default color for non RGB colors.
            Color::Theme(_, _) | Color::Default | Color::Automatic => "#ffffe1".to_string(),
            _ => {
                let rgb_color = Self::rgb_hex_value(self).to_lowercase();
                format!("#{rgb_color}")
            }
        }
    }

    // Get the ARGB hex value for a color. The alpha channel is always FF.
    pub(crate) fn argb_hex_value(self) -> String {
        format!("FF{}", self.rgb_hex_value())
    }

    // Convert the color in a set of "rgb" or "theme/tint" attributes used in
    // color related Style XML elements.
    pub(crate) fn attributes(self) -> Vec<(&'static str, String)> {
        match self {
            Self::Theme(color, shade) => match color {
                // The first 3 columns of colors in the theme palette are
                // different from the others.
                0 => match shade {
                    1 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-4.9989318521683403E-2".to_string()),
                    ],
                    2 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.14999847407452621".to_string()),
                    ],
                    3 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.249977111117893".to_string()),
                    ],
                    4 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.34998626667073579".to_string()),
                    ],
                    5 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.499984740745262".to_string()),
                    ],
                    // The 0 shade is omitted from the attributes.
                    _ => vec![("theme", color.to_string())],
                },
                1 => match shade {
                    1 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.499984740745262".to_string()),
                    ],
                    2 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.34998626667073579".to_string()),
                    ],
                    3 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.249977111117893".to_string()),
                    ],
                    4 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.14999847407452621".to_string()),
                    ],
                    5 => vec![
                        ("theme", color.to_string()),
                        ("tint", "4.9989318521683403E-2".to_string()),
                    ],
                    // The 0 shade is omitted from the attributes.
                    _ => vec![("theme", color.to_string())],
                },
                2 => match shade {
                    1 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-9.9978637043366805E-2".to_string()),
                    ],
                    2 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.249977111117893".to_string()),
                    ],
                    3 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.499984740745262".to_string()),
                    ],
                    4 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.749992370372631".to_string()),
                    ],
                    5 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.89999084444715716".to_string()),
                    ],
                    // The 0 shade is omitted from the attributes.
                    _ => vec![("theme", color.to_string())],
                },
                _ => match shade {
                    1 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.79998168889431442".to_string()),
                    ],
                    2 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.59999389629810485".to_string()),
                    ],
                    3 => vec![
                        ("theme", color.to_string()),
                        ("tint", "0.39997558519241921".to_string()),
                    ],
                    4 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.249977111117893".to_string()),
                    ],
                    5 => vec![
                        ("theme", color.to_string()),
                        ("tint", "-0.499984740745262".to_string()),
                    ],
                    // The 0 shade is omitted from the attributes.
                    _ => vec![("theme", color.to_string())],
                },
            },

            // Handle RGB color.
            _ => vec![("rgb", self.argb_hex_value())],
        }
    }

    // Convert theme colors into the luminance modulation and offset values used
    // in chart theme colors.
    pub(crate) fn chart_scheme(self) -> (String, u32, u32) {
        match self {
            Self::Theme(color, shade) => match color {
                0 => match shade {
                    0 => ("bg1".to_string(), 0, 0),
                    1 => ("bg1".to_string(), 95000, 0),
                    2 => ("bg1".to_string(), 85000, 0),
                    3 => ("bg1".to_string(), 75000, 0),
                    4 => ("bg1".to_string(), 65000, 0),
                    5 => ("bg1".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                1 => match shade {
                    0 => ("tx1".to_string(), 0, 0),
                    1 => ("tx1".to_string(), 50000, 50000),
                    2 => ("tx1".to_string(), 65000, 35000),
                    3 => ("tx1".to_string(), 75000, 25000),
                    4 => ("tx1".to_string(), 85000, 15000),
                    5 => ("tx1".to_string(), 95000, 5000),
                    _ => (String::new(), 0, 0),
                },
                2 => match shade {
                    0 => ("bg2".to_string(), 0, 0),
                    1 => ("bg2".to_string(), 90000, 0),
                    2 => ("bg2".to_string(), 75000, 0),
                    3 => ("bg2".to_string(), 50000, 0),
                    4 => ("bg2".to_string(), 25000, 0),
                    5 => ("bg2".to_string(), 10000, 0),
                    _ => (String::new(), 0, 0),
                },
                3 => match shade {
                    0 => ("tx2".to_string(), 0, 0),
                    1 => ("tx2".to_string(), 20000, 80000),
                    2 => ("tx2".to_string(), 40000, 60000),
                    3 => ("tx2".to_string(), 60000, 40000),
                    4 => ("tx2".to_string(), 75000, 0),
                    5 => ("tx2".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                4 => match shade {
                    0 => ("accent1".to_string(), 0, 0),
                    1 => ("accent1".to_string(), 20000, 80000),
                    2 => ("accent1".to_string(), 40000, 60000),
                    3 => ("accent1".to_string(), 60000, 40000),
                    4 => ("accent1".to_string(), 75000, 0),
                    5 => ("accent1".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                5 => match shade {
                    0 => ("accent2".to_string(), 0, 0),
                    1 => ("accent2".to_string(), 20000, 80000),
                    2 => ("accent2".to_string(), 40000, 60000),
                    3 => ("accent2".to_string(), 60000, 40000),
                    4 => ("accent2".to_string(), 75000, 0),
                    5 => ("accent2".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                6 => match shade {
                    0 => ("accent3".to_string(), 0, 0),
                    1 => ("accent3".to_string(), 20000, 80000),
                    2 => ("accent3".to_string(), 40000, 60000),
                    3 => ("accent3".to_string(), 60000, 40000),
                    4 => ("accent3".to_string(), 75000, 0),
                    5 => ("accent3".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                7 => match shade {
                    0 => ("accent4".to_string(), 0, 0),
                    1 => ("accent4".to_string(), 20000, 80000),
                    2 => ("accent4".to_string(), 40000, 60000),
                    3 => ("accent4".to_string(), 60000, 40000),
                    4 => ("accent4".to_string(), 75000, 0),
                    5 => ("accent4".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                8 => match shade {
                    0 => ("accent5".to_string(), 0, 0),
                    1 => ("accent5".to_string(), 20000, 80000),
                    2 => ("accent5".to_string(), 40000, 60000),
                    3 => ("accent5".to_string(), 60000, 40000),
                    4 => ("accent5".to_string(), 75000, 0),
                    5 => ("accent5".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                9 => match shade {
                    0 => ("accent6".to_string(), 0, 0),
                    1 => ("accent6".to_string(), 20000, 80000),
                    2 => ("accent6".to_string(), 40000, 60000),
                    3 => ("accent6".to_string(), 60000, 40000),
                    4 => ("accent6".to_string(), 75000, 0),
                    5 => ("accent6".to_string(), 50000, 0),
                    _ => (String::new(), 0, 0),
                },
                _ => (String::new(), 0, 0),
            },

            // Handle RGB color with an empty default.
            _ => (String::new(), 0, 0),
        }
    }

    // Check if the RGB and Theme values are in the correct range. Any of the
    // simple enum will be by default.
    #[allow(clippy::unreadable_literal)]
    pub(crate) fn is_valid(self) -> bool {
        match self {
            Color::RGB(color) => {
                if color > 0xFFFFFF {
                    eprintln!(
                        "RGB color '{color:#X}' must be in the the range 0x000000 - 0xFFFFFF."
                    );
                    return false;
                }
                true
            }
            Color::Theme(color, shade) => {
                if color > 9 {
                    eprintln!("Theme color '{color}' must be in the the range 0 - 9.");
                    return false;
                }
                if shade > 5 {
                    eprintln!("Theme shade '{shade}' must be in the the range 0 - 5.");
                    return false;
                }
                true
            }
            _ => true,
        }
    }

    // Check if the color has been set to a non default/automatic color.
    pub(crate) fn is_auto_or_default(self) -> bool {
        self == Color::Automatic || self == Color::Default
    }
}

/// Convert from a u32 RGB value line 0xDAA520 into a [`Color`] enum value.
impl From<u32> for Color {
    fn from(value: u32) -> Color {
        Color::RGB(value)
    }
}

/// Convert from a Html style color string line "#6495ED" into a [`Color`] enum value.
impl From<&str> for Color {
    fn from(value: &str) -> Color {
        let color = if let Some(hex_string) = value.strip_prefix('#') {
            u32::from_str_radix(hex_string, 16)
        } else {
            u32::from_str_radix(value, 16)
        };

        match color {
            Ok(color) => Color::RGB(color),
            Err(_) => {
                eprintln!("Error parsing '{value}' to RGB color.");
                Color::Default
            }
        }
    }
}
