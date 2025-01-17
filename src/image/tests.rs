// Image unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod image_tests {

    use crate::XlsxError;

    use crate::Image;

    #[test]
    fn test_images() {
        let image_test_data = vec![
            // Name, width, height, width_dpi, height_dpi, type.
            ("black_150.jpg", 64, 64, 150.0, 150.0, "jpeg"),
            ("black_300.jpg", 64, 64, 300.0, 300.0, "jpeg"),
            ("black_300.png", 64, 64, 299.9994, 299.9994, "png"),
            ("black_300e.png", 64, 64, 299.9994, 299.9994, "png"),
            ("black_72.jpg", 64, 64, 72.0, 72.0, "jpeg"),
            ("black_72.png", 64, 64, 72.009, 72.009, "png"),
            ("black_72e.png", 64, 64, 72.009, 72.009, "png"),
            ("black_96.jpg", 64, 64, 96.0, 96.0, "jpeg"),
            ("black_96.png", 64, 64, 96.012, 96.012, "png"),
            ("blue.jpg", 23, 23, 96.0, 96.0, "jpeg"),
            ("blue.png", 23, 23, 96.0, 96.0, "png"),
            ("grey.jpg", 99, 69, 96.0, 96.0, "jpeg"),
            ("grey.png", 99, 69, 96.0, 96.0, "png"),
            ("happy.jpg", 423, 563, 96.0, 96.0, "jpeg"),
            ("issue32.png", 115, 115, 96.0, 96.0, "png"),
            ("logo.gif", 200, 80, 96.0, 96.0, "gif"),
            ("logo.jpg", 200, 80, 96.0, 96.0, "jpeg"),
            ("logo.png", 200, 80, 96.0, 96.0, "png"),
            ("mylogo.png", 215, 36, 95.9866, 95.9866, "png"),
            ("red.bmp", 32, 32, 96.0, 96.0, "bmp"),
            ("red.gif", 32, 32, 96.0, 96.0, "gif"),
            ("red.jpg", 32, 32, 96.0, 96.0, "jpeg"),
            ("red.png", 32, 32, 96.0, 96.0, "png"),
            ("red2.png", 32, 32, 96.0, 96.0, "png"),
            ("red_208.png", 208, 49, 96.0, 96.0, "png"),
            ("red_64x20.png", 64, 20, 96.0, 96.0, "png"),
            ("red_readonly.png", 32, 32, 96.0, 96.0, "png"),
            ("train.jpg", 640, 480, 96.0, 96.0, "jpeg"),
            ("watermark.png", 1778, 1003, 329.9968, 329.9968, "png"),
            ("yellow.jpg", 72, 72, 96.0, 96.0, "jpeg"),
            ("yellow.png", 72, 72, 96.0, 96.0, "png"),
            ("zero_dpi.jpg", 11, 16, 96.0, 96.0, "jpeg"),
            (
                "black_150.png",
                64,
                64,
                150.01239999999999,
                150.01239999999999,
                "png",
            ),
            (
                "black_150e.png",
                64,
                64,
                150.01239999999999,
                150.01239999999999,
                "png",
            ),
        ];

        for test_data in image_test_data {
            let (filename, width, height, width_dpi, height_dpi, image_type) = test_data;
            let filename = format!("tests/input/images/{filename}");

            let image = Image::new(&filename).unwrap();
            assert_eq!(width as f64, image.width());
            assert_eq!(height as f64, image.height());
            assert_eq!(width_dpi, image.width_dpi());
            assert_eq!(height_dpi, image.height_dpi());
            assert_eq!(image_type, image.image_type.extension());
        }
    }

    #[test]
    fn unknown_file_format() {
        let filename = "tests/input/images/unknown.img".to_string();

        let image = Image::new(filename);
        assert!(matches!(image, Err(XlsxError::UnknownImageType)));
    }

    #[test]
    fn invalid_file_format() {
        let filename = "tests/input/images/no_dimensions.png".to_string();

        let image = Image::new(filename);
        assert!(matches!(image, Err(XlsxError::ImageDimensionError)));
    }
}
