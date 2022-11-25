// image - A module for handling Excel image files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::fs::File;
use std::io::BufReader;
use std::io::Read;
use std::path::Path;

use crate::XlsxError;

#[derive(Clone, Debug)]
/// TODO
pub struct Image {
    height: u32,
    width: u32,
    x_dpi: f64,
    y_dpi: f64,
    read_dpi: bool,
    image_type: XlsxImageType,
    alt_text: String,
}

impl Image {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// TODO
    pub fn new(filename: &str) -> Result<Image, XlsxError> {
        let path = Path::new(&filename);
        let file = File::open(path).unwrap();
        let mut reader = BufReader::new(file);
        let mut image_data: Vec<u8> = vec![];
        reader.read_to_end(&mut image_data).unwrap();

        let mut image = Image {
            height: 0,
            width: 0,
            x_dpi: 96.0,
            y_dpi: 96.0,
            read_dpi: false,
            image_type: XlsxImageType::Unknown,
            alt_text: "".to_string(),
        };

        Self::process_image(&mut image, &image_data);

        Ok(image)
    }

    /// TODO
    pub fn width(&self) -> u32 {
        self.width
    }

    /// TODO
    pub fn height(&self) -> u32 {
        self.height
    }

    /// TODO
    pub fn x_dpi(&self) -> f64 {
        self.x_dpi
    }

    /// TODO
    pub fn y_dpi(&self) -> f64 {
        self.y_dpi
    }

    /// TODO
    pub fn image_type(&self) -> String {
        self.image_type.value()
    }

    /// TODO
    pub fn set_alt_text(mut self, alt_text: &str) -> Image {
        self.alt_text = alt_text.to_string();
        self
    }

    // -----------------------------------------------------------------------
    // Internal methods.
    // -----------------------------------------------------------------------

    // Extract type and width and height information from an image file.
    fn process_image(&mut self, data: &[u8]) {
        let png_marker = &data[1..4];
        let jpg_marker = unpack_u16_from_be_bytes(&data, 0);
        let bmp_marker = &data[0..2];
        let jif_marker = &data[0..4];

        if png_marker == "PNG".as_bytes() {
            self.process_png(&data);
        } else if jpg_marker == 0xFFD8 {
            self.process_jpg(&data);
        } else if bmp_marker == "BM".as_bytes() {
            self.process_bmp(&data);
        } else if jif_marker == "GIF8".as_bytes() {
            self.process_jif(&data);
        }
    }

    // Extract width and height information from a PNG file.
    fn process_png(&mut self, data: &[u8]) {
        let mut offset: usize = 8;
        let mut width: u32 = 0;
        let mut height: u32 = 0;
        let mut x_dpi: f64 = 96.0;
        let mut y_dpi: f64 = 96.0;
        let data_length = data.len();

        // Search through the image data to read the height and width in the
        // IHDR element. Also read the DPI in the pHYs element, if present.
        while offset < data_length {
            let marker = &data[offset + 4..offset + 8];
            let length = unpack_u32_from_be_bytes(data, offset);

            // Read the image dimensions.
            if marker == "IHDR".as_bytes() {
                width = unpack_u32_from_be_bytes(data, offset + 8);
                height = unpack_u32_from_be_bytes(data, offset + 12);
            }

            // Read the image DPI values.
            if marker == "pHYs".as_bytes() {
                let units = &data[offset + 16];
                let x_density = unpack_u32_from_be_bytes(data, offset + 8);
                let y_density = unpack_u32_from_be_bytes(data, offset + 12);

                if *units == 1 {
                    x_dpi = x_density as f64 * 0.0254;
                    y_dpi = y_density as f64 * 0.0254;
                    self.read_dpi = true;
                }
            }

            if marker == "IEND".as_bytes() {
                break;
            }

            offset = offset + length as usize + 12;
        }

        self.width = width;
        self.height = height;
        self.x_dpi = x_dpi;
        self.y_dpi = y_dpi;
        self.image_type = XlsxImageType::PNG;
    }

    // Extract width and height information from a PNG file.
    fn process_jpg(&mut self, data: &[u8]) {
        let mut offset: usize = 2;
        let mut height: u32 = 0;
        let mut width: u32 = 0;
        let mut x_dpi: f64 = 96.0;
        let mut y_dpi: f64 = 96.0;
        let data_length = data.len();

        // Search through the image data to read the height and width in the
        // IHDR element. Also read the DPI in the pHYs element, if present.
        while offset < data_length {
            let marker = unpack_u16_from_be_bytes(data, offset);
            let length = unpack_u16_from_be_bytes(data, offset + 2);

            // Read the height and width in the 0xFFCn elements (except C4, C8
            // and CC which aren't SOF markers).
            if (marker & 0xFFF0) == 0xFFC0
                && marker != 0xFFC4
                && marker != 0xFFC8
                && marker != 0xFFCC
            {
                height = unpack_u16_from_be_bytes(data, offset + 5) as u32;
                width = unpack_u16_from_be_bytes(data, offset + 7) as u32;
            }

            // Read the DPI in the 0xFFE0 element.
            if marker == 0xFFE0 {
                let units = &data[offset + 11];
                let x_density = unpack_u16_from_be_bytes(data, offset + 12);
                let y_density = unpack_u16_from_be_bytes(data, offset + 14);

                if *units == 1 {
                    x_dpi = x_density as f64;
                    y_dpi = y_density as f64;
                }

                if *units == 2 {
                    x_dpi = x_density as f64 * 2.54;
                    y_dpi = y_density as f64 * 2.54;
                    self.read_dpi = true;
                }

                // Workaround for incorrect dpi.
                if x_dpi == 0.0 || x_dpi == 1.0 {
                    x_dpi = 96.0
                }
                if y_dpi == 0.0 || y_dpi == 1.0 {
                    y_dpi = 96.0
                }
            }

            if marker == 0xFFDA {
                break;
            }

            offset = offset + length as usize + 2;
        }

        self.width = width;
        self.height = height;
        self.x_dpi = x_dpi;
        self.y_dpi = y_dpi;
        self.image_type = XlsxImageType::JPG;
    }

    // Extract width and height information from a BMP file.
    fn process_bmp(&mut self, data: &[u8]) {
        let x_dpi: f64 = 96.0;
        let y_dpi: f64 = 96.0;

        let width = unpack_u32_from_le_bytes(data, 18);
        let height = unpack_u32_from_le_bytes(data, 22);

        self.width = width;
        self.height = height;
        self.x_dpi = x_dpi;
        self.y_dpi = y_dpi;
        self.image_type = XlsxImageType::BMP;
    }

    // Extract width and height information from a GIF file.
    fn process_jif(&mut self, data: &[u8]) {
        let width = unpack_u16_from_le_bytes(data, 6) as u32;
        let height = unpack_u16_from_le_bytes(data, 8) as u32;

        self.width = width;
        self.height = height;
        self.x_dpi = 96.0;
        self.y_dpi = 96.0;
        self.image_type = XlsxImageType::GIF;
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------
#[derive(Clone, Debug)]
pub(crate) enum XlsxImageType {
    Unknown,
    PNG,
    JPG,
    GIF,
    BMP,
}

impl XlsxImageType {
    pub(crate) fn value(&self) -> String {
        match self {
            XlsxImageType::Unknown => "Unknown image type".to_string(),
            XlsxImageType::PNG => "PNG".to_string(),
            XlsxImageType::JPG => "JPG".to_string(),
            XlsxImageType::GIF => "GIF".to_string(),
            XlsxImageType::BMP => "BMP".to_string(),
        }
    }
}

// Some helper functions to extract 2 and 4 byte integers from image data.
fn unpack_u16_from_be_bytes(data: &[u8], offset: usize) -> u16 {
    u16::from_be_bytes(data[offset..offset + 2].try_into().unwrap())
}

fn unpack_u16_from_le_bytes(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
}

fn unpack_u32_from_be_bytes(data: &[u8], offset: usize) -> u32 {
    u32::from_be_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn unpack_u32_from_le_bytes(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::Image;

    #[test]
    fn test_images() {
        let image_test_data = vec![
            // Name, width, height, x_dpi, y_dpi, type.
            ("black_150.jpg", 64, 64, 150.0, 150.0, "JPG"),
            ("black_300.jpg", 64, 64, 300.0, 300.0, "JPG"),
            ("black_300.png", 64, 64, 299.9994, 299.9994, "PNG"),
            ("black_300e.png", 64, 64, 299.9994, 299.9994, "PNG"),
            ("black_72.jpg", 64, 64, 72.0, 72.0, "JPG"),
            ("black_72.png", 64, 64, 72.009, 72.009, "PNG"),
            ("black_72e.png", 64, 64, 72.009, 72.009, "PNG"),
            ("black_96.jpg", 64, 64, 96.0, 96.0, "JPG"),
            ("black_96.png", 64, 64, 96.012, 96.012, "PNG"),
            ("blue.jpg", 23, 23, 96.0, 96.0, "JPG"),
            ("blue.png", 23, 23, 96.0, 96.0, "PNG"),
            ("grey.jpg", 99, 69, 96.0, 96.0, "JPG"),
            ("grey.png", 99, 69, 96.0, 96.0, "PNG"),
            ("happy.jpg", 423, 563, 96.0, 96.0, "JPG"),
            ("issue32.png", 115, 115, 96.0, 96.0, "PNG"),
            ("logo.gif", 200, 80, 96.0, 96.0, "GIF"),
            ("logo.jpg", 200, 80, 96.0, 96.0, "JPG"),
            ("logo.png", 200, 80, 96.0, 96.0, "PNG"),
            ("mylogo.png", 215, 36, 95.9866, 95.9866, "PNG"),
            ("red.bmp", 32, 32, 96.0, 96.0, "BMP"),
            ("red.gif", 32, 32, 96.0, 96.0, "GIF"),
            ("red.jpg", 32, 32, 96.0, 96.0, "JPG"),
            ("red.png", 32, 32, 96.0, 96.0, "PNG"),
            ("red2.png", 32, 32, 96.0, 96.0, "PNG"),
            ("red_208.png", 208, 49, 96.0, 96.0, "PNG"),
            ("red_64x20.png", 64, 20, 96.0, 96.0, "PNG"),
            ("red_readonly.png", 32, 32, 96.0, 96.0, "PNG"),
            ("train.jpg", 640, 480, 96.0, 96.0, "JPG"),
            ("watermark.png", 1778, 1003, 329.9968, 329.9968, "PNG"),
            ("yellow.jpg", 72, 72, 96.0, 96.0, "JPG"),
            ("yellow.png", 72, 72, 96.0, 96.0, "PNG"),
            ("zero_dpi.jpg", 11, 16, 96.0, 96.0, "JPG"),
            (
                "black_150.png",
                64,
                64,
                150.01239999999999,
                150.01239999999999,
                "PNG",
            ),
            (
                "black_150e.png",
                64,
                64,
                150.01239999999999,
                150.01239999999999,
                "PNG",
            ),
        ];

        for test_data in image_test_data {
            let (filename, width, height, x_dpi, y_dpi, image_type) = test_data;
            let filename = format!("tests/input/images/{filename}");

            let image = Image::new(&filename).unwrap();
            assert_eq!(width, image.width());
            assert_eq!(height, image.height());
            assert_eq!(x_dpi, image.x_dpi());
            assert_eq!(y_dpi, image.y_dpi());
            assert_eq!(image_type, image.image_type());
        }
    }
}
