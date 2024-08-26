// image - A module for handling Excel image files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::collections::hash_map::DefaultHasher;
use std::fs::File;
use std::hash::{Hash, Hasher};
use std::io::BufReader;
use std::io::Read;
use std::path::Path;
use std::path::PathBuf;

use crate::drawing::{DrawingObject, DrawingType};
use crate::vml::VmlInfo;
use crate::{Url, XlsxError};

#[derive(Clone, Debug)]
/// The `Image` struct is used to create an object to represent an image that
/// can be inserted into a worksheet.
///
/// ```rust
/// # // This code is available in examples/doc_image.rs
/// #
/// use rust_xlsxwriter::{Image, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Create a new image object.
///     let image = Image::new("examples/rust_logo.png")?;
///
///     // Insert the image.
///     worksheet.insert_image(1, 2, &image)?;
///
///     // Save the file to disk.
///     workbook.save("image.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/image_intro.png">
///
pub struct Image {
    height: f64,
    width: f64,
    width_dpi: f64,
    height_dpi: f64,
    scale_width: f64,
    scale_height: f64,
    has_default_dpi: bool,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) image_type: XlsxImageType,
    pub(crate) name: String,
    pub(crate) alt_text: String,
    pub(crate) vml_name: String,
    pub(crate) header_position: HeaderImagePosition,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) is_header: bool,
    pub(crate) decorative: bool,
    pub(crate) hash: String,
    pub(crate) data: Vec<u8>,
    pub(crate) drawing_type: DrawingType,
    pub(crate) url: Option<Url>,
}

impl Image {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Image object from an image file.
    ///
    /// Create an Image object from a path to an image file. The image can then
    /// be inserted into a worksheet.
    ///
    /// The supported image formats are:
    ///
    /// - PNG
    /// - JPG
    /// - GIF: The image can be an animated gif in more recent versions of
    ///   Excel.
    /// - BMP: BMP images are only supported for backward compatibility. In
    ///   general it is best to avoid BMP images since they are not compressed.
    ///   If used, BMP images must be 24 bit, true color, bitmaps.
    ///
    /// EMF and WMF file formats will be supported in an upcoming version of the
    /// library.
    ///
    /// **NOTE on SVG files**: Excel doesn't directly support SVG files in the
    /// same way as other image file formats. It allows SVG to be inserted into
    /// a worksheet but converts them to, and displays them as, PNG files. It
    /// stores the original SVG image in the file so the original format can be
    /// retrieved. This removes the file size and resolution advantage of using
    /// SVG files. As such SVG files are not supported by `rust_xlsxwriter`
    /// since a conversion to the PNG format would be required and that format
    /// is already supported.
    ///
    /// # Parameters
    ///
    /// - `path`: The path of the image file to read e as a `&str` or as a
    ///   [`std::path`] `Path` or `PathBuf` instance.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::UnknownImageType`] - Unknown image type. The supported
    ///   image formats are PNG, JPG, GIF and BMP.
    /// - [`XlsxError::ImageDimensionError`] - Image has 0 width or height, or
    ///   the dimensions couldn't be read.
    ///
    /// # Examples
    ///
    /// The following example demonstrates creating a new Image object and
    /// adding it to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_image.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new image object.
    ///     let image = Image::new("examples/rust_logo.png")?;
    ///
    ///     // Insert the image.
    ///     worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/image_intro.png">
    ///
    pub fn new<P: AsRef<Path>>(path: P) -> Result<Image, XlsxError> {
        let mut path_buf = PathBuf::new();
        path_buf.push(path);

        let vml_name = match path_buf.file_stem() {
            Some(file_stem) => file_stem.to_string_lossy().to_string(),
            None => "image".to_string(),
        };

        let file = File::open(path_buf)?;
        let mut reader = BufReader::new(file);
        let mut data = vec![];
        reader.read_to_end(&mut data)?;

        let mut image = Self::new_from_buffer(&data)?;
        image.vml_name = vml_name;

        Ok(image)
    }

    /// Create an Image object from a u8 buffer. The image can then be inserted
    /// into a worksheet.
    ///
    /// This method is similar to [`Image::new`], see above, except the image
    /// data can be in a buffer instead of a file path.
    ///
    /// # Parameters
    ///
    /// - `buffer`: The image data as a u8 array or vector.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::UnknownImageType`] - Unknown image type. The supported
    ///   image formats are PNG, JPG, GIF and BMP.
    /// - [`XlsxError::ImageDimensionError`] - Image has 0 width or height, or
    ///   the dimensions couldn't be read.
    ///
    /// # Examples
    ///
    /// This example shows how to create an image object from a u8 buffer.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_new_from_buffer.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Create a new image object.
    /// #     let buf: [u8; 200] = [
    /// #         0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
    /// #         0x52, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20, 0x08, 0x02, 0x00, 0x00, 0x00, 0xfc,
    /// #         0x18, 0xed, 0xa3, 0x00, 0x00, 0x00, 0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xae, 0xce, 0x1c,
    /// #         0xe9, 0x00, 0x00, 0x00, 0x04, 0x67, 0x41, 0x4d, 0x41, 0x00, 0x00, 0xb1, 0x8f, 0x0b, 0xfc,
    /// #         0x61, 0x05, 0x00, 0x00, 0x00, 0x20, 0x63, 0x48, 0x52, 0x4d, 0x00, 0x00, 0x7a, 0x26, 0x00,
    /// #         0x00, 0x80, 0x84, 0x00, 0x00, 0xfa, 0x00, 0x00, 0x00, 0x80, 0xe8, 0x00, 0x00, 0x75, 0x30,
    /// #         0x00, 0x00, 0xea, 0x60, 0x00, 0x00, 0x3a, 0x98, 0x00, 0x00, 0x17, 0x70, 0x9c, 0xba, 0x51,
    /// #         0x3c, 0x00, 0x00, 0x00, 0x46, 0x49, 0x44, 0x41, 0x54, 0x48, 0x4b, 0x63, 0xfc, 0xcf, 0x40,
    /// #         0x63, 0x00, 0xb4, 0x80, 0xa6, 0x88, 0xb6, 0xa6, 0x83, 0x82, 0x87, 0xa6, 0xce, 0x1f, 0xb5,
    /// #         0x80, 0x98, 0xe0, 0x1d, 0x8d, 0x03, 0x82, 0xa1, 0x34, 0x1a, 0x44, 0xa3, 0x41, 0x44, 0x30,
    /// #         0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45, 0xa3, 0x41, 0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d,
    /// #         0x45, 0xa3, 0x41, 0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45, 0x03, 0x1f, 0x44, 0x00,
    /// #         0xaa, 0x35, 0xdd, 0x4e, 0xe6, 0xd5, 0xa1, 0x22, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    /// #         0x44, 0xae, 0x42, 0x60, 0x82,
    /// #     ];
    /// #
    ///     // Create a new image object from a u8 buffer.
    ///     let image = Image::new_from_buffer(&buf)?;
    ///
    ///     // Insert the image.
    ///     worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/image_new_from_buffer.png">
    ///
    pub fn new_from_buffer(buffer: &[u8]) -> Result<Image, XlsxError> {
        let mut image = Image {
            height: 0.0,
            width: 0.0,
            width_dpi: 96.0,
            height_dpi: 96.0,
            scale_width: 1.0,
            scale_height: 1.0,
            x_offset: 0,
            y_offset: 0,
            has_default_dpi: true,
            image_type: XlsxImageType::Unknown,
            name: String::new(),
            alt_text: String::new(),
            vml_name: "image".to_string(),
            header_position: HeaderImagePosition::Center,
            object_movement: ObjectMovement::MoveButDontSizeWithCells,
            is_header: true,
            decorative: false,
            hash: String::new(),
            data: buffer.to_vec(),
            drawing_type: DrawingType::Image,
            url: None,
        };

        Self::process_image(&mut image)?;

        Ok(image)
    }

    /// Set the width of the chart.
    ///
    /// Set the displayed width of the image in pixels. As with Excel this sets
    /// a logical size for the image, it doesn't rescale the actual image. This
    /// allows the user to get back the original unscaled image.
    ///
    /// **Note for macOS Excel users**: the width shown on Excel for macOS can
    /// be different from the width on Windows. This is an Excel issue and not a
    /// `rust_xlsxwriter` issue.
    ///
    /// # Parameters
    ///
    /// - `width`: The logical image width in pixels.
    ///
    /// # Examples
    ///
    /// This example shows how to create an image object and use it to insert
    /// the image into a worksheet. The image in this case is scaled by setting
    /// the height and width.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_set_width.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new image object and set the image logical sizes.
    ///     let image = Image::new("examples/rust_logo.png")?
    ///         .set_height(80)
    ///         .set_width(80);
    ///
    ///     // Insert the image.
    ///     worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/image_set_scale_width.png">
    ///
    pub fn set_width(mut self, width: u32) -> Image {
        if width == 0 {
            return self;
        }

        // Set the scale width rather than the actual height.
        self.scale_width = f64::from(width) / self.width;
        self
    }

    /// Set the height of the image.
    ///
    /// Set the displayed height of the image in pixels. As with Excel this sets
    /// a logical size for the image, it doesn't rescale the actual image. This
    /// allows the user to get back the original unscaled image. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// - `height`: The logical image height in pixels.
    ///
    pub fn set_height(mut self, height: u32) -> Image {
        if height == 0 {
            return self;
        }

        // Set the scale height rather than the actual height.
        self.scale_height = f64::from(height) / self.height;
        self
    }

    /// Set the height scale for the image.
    ///
    /// Set the height scale for the image relative to 1.0 (i.e. 100%). As with Excel
    /// this sets a logical scale for the image, it doesn't rescale the actual
    /// image. This allows the user to get back the original unscaled image.
    ///
    /// **Note for macOS Excel users**: the scale shown on Excel for macOS is
    /// different from the scale on Windows. This is an Excel issue and not a
    /// `rust_xlsxwriter` issue.
    ///
    /// # Parameters
    ///
    /// - `scale`: The scale ratio.
    ///
    /// # Examples
    ///
    /// This example shows how to create an image object and use it to insert
    /// the image into a worksheet. The image in this case is scaled.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_set_scale_width.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new image object and set the image scale.
    ///     let image = Image::new("examples/rust_logo.png")?
    ///         .set_scale_height(0.75)
    ///         .set_scale_width(0.75);
    ///
    ///     // Insert the image.
    ///     worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/image_set_scale_width.png">
    ///
    pub fn set_scale_height(mut self, scale: f64) -> Image {
        if scale <= 0.0 {
            return self;
        }

        self.scale_height = scale;
        self
    }

    /// Set the width scale for the image.
    ///
    /// Set the width scale for the image relative to 1.0 (i.e. 100%). See the
    /// [`Image::set_scale_height()`] method for details.
    ///
    /// # Parameters
    ///
    /// - `scale`: The scale ratio.
    ///
    pub fn set_scale_width(mut self, scale: f64) -> Image {
        if scale <= 0.0 {
            return self;
        }

        self.scale_width = scale;
        self
    }

    /// Set the width and height scale to achieve a specific size.
    ///
    /// Calculate and set the horizontal and vertical scales for an image in
    /// order to display it at a fixed width and height in a worksheet. This is
    /// most commonly used to scale an image so that it fits within a cell or a
    /// specific region in a worksheet. The scaling calculation takes into
    /// account the DPI of the image in the same way that Excel does.
    ///
    /// There are two options, which are controlled by the `keep_aspect_ratio`
    /// parameter. The image can be scaled vertically and horizontally to give
    /// the specified with and height or the aspect ratio of the image can be
    /// maintained so that the image is scaled to the lesser of the horizontal
    /// or vertical sizes. See the example below.
    ///
    /// See also the
    /// [`Worksheet::insert_image_fit_to_cell()`](crate::Worksheet::insert_image_fit_to_cell)
    /// method.
    ///
    /// # Parameters
    ///
    /// - `width`: The target width in pixels to scale the image to.
    /// - `height`: The target height in pixels to scale the image to.
    /// - `keep_aspect_ratio`: Boolean value to maintain the aspect ratio of
    ///   the image if `true` or scale independently in the horizontal and
    ///   vertical directions if `false`.
    ///
    /// Note: the `width` and `height` can mainly be considered as pixel sizes.
    /// However, f64 values are allowed for cases where a fractional size is
    /// required
    ///
    /// # Examples
    ///
    /// An example of scaling images to a fixed width and height. See also the
    /// `worksheet.insert_image_fit_to_cell()` method.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_set_scale_to_size.rs
    /// #
    /// # use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let center = Format::new().set_align(FormatAlign::VerticalCenter);
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Widen the first column to make the text clearer.
    ///     worksheet.set_column_width(0, 30)?;
    ///
    ///     // Set larger cells to accommodate the images.
    ///     worksheet.set_column_width_pixels(1, 200)?;
    ///     worksheet.set_row_height_pixels(0, 140)?;
    ///     worksheet.set_row_height_pixels(2, 140)?;
    ///     worksheet.set_row_height_pixels(4, 140)?;
    ///
    ///     // Create a new image object.
    ///     let mut image = Image::new("examples/rust_logo.png")?;
    ///
    ///     // Insert the image as standard, without scaling.
    ///     worksheet.write_with_format(0, 0, "Unscaled image inserted into cell:", &center)?;
    ///     worksheet.insert_image(0, 1, &image)?;
    ///
    ///     // Scale the image to fit the entire cell.
    ///     image = image.set_scale_to_size(200, 140, false);
    ///     worksheet.write_with_format(2, 0, "Image scaled to fit cell:", &center)?;
    ///     worksheet.insert_image(2, 1, &image)?;
    ///
    ///     // Scale the image to fit the defined size region while maintaining the
    ///     // aspect ratio. In this case it is scaled to the smaller of the width or
    ///     // height scales.
    ///     image = image.set_scale_to_size(200, 140, true);
    ///     worksheet.write_with_format(4, 0, "Image scaled with a fixed aspect ratio:", &center)?;
    ///     worksheet.insert_image(4, 1, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/image_set_scale_to_size.png">
    ///
    ///
    pub fn set_scale_to_size<T>(mut self, width: T, height: T, keep_aspect_ratio: bool) -> Image
    where
        T: Into<f64> + Copy,
    {
        if width.into() == 0.0 || height.into() == 0.0 {
            return self;
        }

        let mut scale_width = (width.into() / self.width()) * (self.width_dpi() / 96.0);
        let mut scale_height = (height.into() / self.height()) * (self.height_dpi() / 96.0);

        if keep_aspect_ratio {
            if scale_width < scale_height {
                scale_height = scale_width;
            } else {
                scale_width = scale_height;
            }
        }

        self = self.set_scale_width(scale_width);
        self = self.set_scale_height(scale_height);

        self
    }

    /// Set the alt text for the image.
    ///
    /// Set the alt text for the image to help accessibility. The alt text is
    /// used with screen readers to help people with visual disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the image.
    ///
    /// # Examples
    ///
    /// This example shows how to create an image object and set the alternative
    /// text to help accessibility.
    ///
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_image_set_alt_text.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #      let mut workbook = Workbook::new();
    /// #
    /// #    // Add a worksheet to the workbook.
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    ///    // Create a new image object and set the alternative text.
    ///    let image = Image::new("examples/rust_logo.png")?.set_alt_text(
    ///        "A circular logo with gear teeth on the outside \
    ///        and a large letter R on the inside.\n\n\
    ///        The logo of the Rust programming language.",
    ///    );
    /// #
    /// #    // Insert the image.
    /// #    worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #    // Save the file to disk.
    /// #      workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Alt text dialog in Excel:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/image_set_alt_text.png">
    ///
    pub fn set_alt_text(mut self, alt_text: impl Into<String>) -> Image {
        let alt_text = alt_text.into();
        if alt_text.chars().count() > 255 {
            eprintln!("Alternative text is greater than Excel's limit of 255 characters.");
            return self;
        }

        self.alt_text = alt_text;
        self
    }

    /// Mark an image as decorative.
    ///
    /// Not all images need an alt text description. Some images may contain
    /// little or no useful visual information, for example a simple shape with
    /// color used to divide sections. Such images can be marked as "decorative"
    /// so that screen readers can inform the users that they don't contain
    /// important information.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// This example shows how to create an image object and set the decorative
    /// property to indicate the it doesn't contain useful visual information.
    /// This is used to improve the accessibility of visual elements.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_set_decorative.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #    // Add a worksheet to the workbook.
    /// #    let worksheet = workbook.add_worksheet();
    /// #
    ///    // Create a new image object.
    ///    let image = Image::new("examples/rust_logo.png")?.set_decorative(true);
    /// #
    /// #    // Insert the image.
    /// #    worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #    // Save the file to disk.
    /// #    workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Alt text dialog in Excel:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/image_set_decorative.png">
    ///
    pub fn set_decorative(mut self, enable: bool) -> Image {
        self.decorative = enable;
        self
    }

    /// Set the object movement options for a worksheet image.
    ///
    /// Set the option to define how an image will behave in Excel if the cells
    /// under the image are moved, deleted, or have their size changed. In Excel
    /// the options are:
    ///
    /// 1. Move and size with cells.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/object_movement.png">
    ///
    /// These values are defined in the [`ObjectMovement`] enum.
    ///
    /// The [`ObjectMovement`] enum also provides an additional option to
    /// "Move and size with cells - after the image is inserted" to allow images
    /// to be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal image position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// - `option`: An image/object positioning behavior defined by the
    ///   [`ObjectMovement`] enum.
    ///
    /// # Examples
    ///
    /// This example shows how to create an image object and set the option to
    /// control how it behaves when the cells underneath it are changed.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_set_object_movement.rs
    /// #
    /// # use rust_xlsxwriter::{Image, Workbook, XlsxError, ObjectMovement};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new image and set the object movement/positioning options.
    ///     let image = Image::new("examples/rust_logo.png")?
    ///         .set_object_movement(ObjectMovement::MoveButDontSizeWithCells);
    ///
    ///     // Insert the image.
    ///     worksheet.insert_image(1, 2, &image)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("image.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/image_set_object_movement.png">
    ///
    pub fn set_object_movement(mut self, option: ObjectMovement) -> Image {
        self.object_movement = option;
        self
    }

    /// Set a Url/Hyperlink for an image.
    ///
    /// Set a Url/Hyperlink for an image so that when the user clicks on it they
    /// are redirected to an internal or external location.
    ///
    /// See [`Url`] for an explanation of the URIs supported by Excel and for
    /// other options that can be set.
    ///
    /// # Parameters
    ///
    /// - `link`: The url/hyperlink associate with the image as a string or
    ///   [`Url`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::MaxUrlLengthExceeded`] - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - [`XlsxError::UnknownUrlType`] - The URL has an unknown URI type. See
    ///   [`Worksheet::write_url()`](crate::Worksheet::write_url).
    /// - [`XlsxError::ParameterError`] - URL mouseover tool tip exceeds Excel's
    ///   limit of 255 characters.
    ///
    pub fn set_url(mut self, link: impl Into<Url>) -> Result<Image, XlsxError> {
        let mut url = link.into();
        url.initialize()?;

        self.url = Some(url);

        Ok(self)
    }

    /// Get the width of the image used for the size calculations in Excel.
    ///
    /// Note, this gets the actual pixel width of the image and not the
    /// logical/scaled width set via [`Image::set_width()`].
    ///
    /// # Examples
    ///
    /// This example shows how to get some of the properties of an Image that
    /// will be used in an Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_image_dimensions.rs
    /// #
    /// # use rust_xlsxwriter::{Image, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let image = Image::new("examples/rust_logo.png")?;
    ///
    ///     assert_eq!(106.0, image.width());
    ///     assert_eq!(106.0, image.height());
    ///     assert_eq!(96.0, image.width_dpi());
    ///     assert_eq!(96.0, image.height_dpi());
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn width(&self) -> f64 {
        self.width
    }

    /// Get the height of the image used for the size calculations in Excel. See
    /// the example above.
    ///
    /// Note, this gets the actual pixel height of the image and not the
    /// logical/scaled height set via [`Image::set_height()`].
    ///
    pub fn height(&self) -> f64 {
        self.height
    }

    /// Get the width/horizontal DPI of the image used for the size calculations
    /// in Excel. See the example above.
    ///
    /// Excel assumes a default image DPI of 96.0 and scales all other DPIs
    /// relative to that.
    ///
    pub fn width_dpi(&self) -> f64 {
        self.width_dpi
    }

    /// Get the height/vertical DPI of the image used for the size calculations
    /// in Excel. See the example above.
    ///
    /// Excel assumes a default image DPI of 96.0 and scales all other DPIs
    /// relative to that.
    ///
    pub fn height_dpi(&self) -> f64 {
        self.height_dpi
    }

    /// Set an internal name used for header/footer images.
    ///
    /// This method sets an internal image name used by header/footer VML. It is
    /// mainly used for completeness in testing. It isn't useful to the end user.
    ///
    /// # Parameters
    ///
    /// `name` - The VML object name/description.
    ///
    #[doc(hidden)]
    pub fn set_vml_name(mut self, name: impl Into<String>) -> Image {
        self.vml_name = name.into();
        self
    }

    // Header images are stored in a vmlDrawing file. We create a struct
    // to store the required image information in that format.
    pub(crate) fn vml_info(&self) -> VmlInfo {
        VmlInfo {
            width: self.vml_width(),
            height: self.vml_height(),
            text: self.vml_name(),
            header_position: self.vml_position(),
            is_scaled: self.is_scaled(),
            ..Default::default()
        }
    }

    // Get the image width as used by header/footer VML.
    fn vml_width(&self) -> f64 {
        // Scale the image dimension relative to 96dpi.
        self.width * 96.0 / self.width_dpi * self.scale_width
    }

    // Get the image height as used by header/footer VML.
    fn vml_height(&self) -> f64 {
        // Scale the image dimension relative to 96dpi.
        self.height * 96.0 / self.height_dpi * self.scale_height
    }

    // Get the image short name as used by header/footer VML.
    fn vml_name(&self) -> String {
        self.vml_name.clone()
    }

    // Check if the image scale has changed. Mainly used by header/footer VML.
    pub(crate) fn is_scaled(&self) -> bool {
        self.scale_height != 1.0 || self.scale_width != 1.0
    }

    // Get the image position string as used by header/footer VML.
    fn vml_position(&self) -> String {
        if self.is_header {
            match self.header_position {
                HeaderImagePosition::Left => "LH".to_string(),
                HeaderImagePosition::Right => "RH".to_string(),
                HeaderImagePosition::Center => "CH".to_string(),
            }
        } else {
            match self.header_position {
                HeaderImagePosition::Left => "LF".to_string(),
                HeaderImagePosition::Right => "RF".to_string(),
                HeaderImagePosition::Center => "CF".to_string(),
            }
        }
    }

    // -----------------------------------------------------------------------
    // Internal methods.
    // -----------------------------------------------------------------------

    // Extract type and width and height information from an image file.
    fn process_image(&mut self) -> Result<(), XlsxError> {
        let data = self.data.clone();

        let png_marker = &data[1..4];
        let jpg_marker = unpack_u16_from_be_bytes(&data, 0);
        let bmp_marker = &data[0..2];
        let gif_marker = &data[0..4];

        if png_marker == "PNG".as_bytes() {
            self.process_png(&data);
        } else if jpg_marker == 0xFFD8 {
            self.process_jpg(&data);
        } else if bmp_marker == "BM".as_bytes() {
            self.process_bmp(&data);
        } else if gif_marker == "GIF8".as_bytes() {
            self.process_gif(&data);
        }

        // Check that we read a valid image.
        if let XlsxImageType::Unknown = self.image_type {
            return Err(XlsxError::UnknownImageType);
        }

        // Check that we read a the image dimensions.
        if self.width == 0.0 || self.height == 0.0 {
            return Err(XlsxError::ImageDimensionError);
        }

        // Set a hash for the image to allow removal of duplicates.
        let mut hasher = DefaultHasher::new();
        data.hash(&mut hasher);
        self.hash = hasher.finish().to_string();

        Ok(())
    }

    // Extract width and height information from a PNG file.
    fn process_png(&mut self, data: &[u8]) {
        let mut offset: usize = 8;
        let mut width: u32 = 0;
        let mut height: u32 = 0;
        let mut width_dpi: f64 = 96.0;
        let mut height_dpi: f64 = 96.0;
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
                    width_dpi = f64::from(x_density) * 0.0254;
                    height_dpi = f64::from(y_density) * 0.0254;
                    self.has_default_dpi = false;
                }
            }

            if marker == "IEND".as_bytes() {
                break;
            }

            offset = offset + length as usize + 12;
        }

        self.width = f64::from(width);
        self.height = f64::from(height);
        self.width_dpi = width_dpi;
        self.height_dpi = height_dpi;
        self.image_type = XlsxImageType::Png;
    }

    // Extract width and height information from a PNG file.
    fn process_jpg(&mut self, data: &[u8]) {
        let mut offset: usize = 2;
        let mut height: u32 = 0;
        let mut width: u32 = 0;
        let mut width_dpi: f64 = 96.0;
        let mut height_dpi: f64 = 96.0;
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
                height = u32::from(unpack_u16_from_be_bytes(data, offset + 5));
                width = u32::from(unpack_u16_from_be_bytes(data, offset + 7));
            }

            // Read the DPI in the 0xFFE0 element.
            if marker == 0xFFE0 {
                let units = &data[offset + 11];
                let x_density = unpack_u16_from_be_bytes(data, offset + 12);
                let y_density = unpack_u16_from_be_bytes(data, offset + 14);

                if *units == 1 {
                    width_dpi = f64::from(x_density);
                    height_dpi = f64::from(y_density);
                }

                if *units == 2 {
                    width_dpi = f64::from(x_density) * 2.54;
                    height_dpi = f64::from(y_density) * 2.54;
                    self.has_default_dpi = false;
                }

                // Workaround for incorrect dpi.
                if width_dpi == 0.0 || width_dpi == 1.0 {
                    width_dpi = 96.0;
                }
                if height_dpi == 0.0 || height_dpi == 1.0 {
                    height_dpi = 96.0;
                }
            }

            if marker == 0xFFDA {
                break;
            }

            offset = offset + length as usize + 2;
        }

        self.width = f64::from(width);
        self.height = f64::from(height);
        self.width_dpi = width_dpi;
        self.height_dpi = height_dpi;
        self.image_type = XlsxImageType::Jpg;
    }

    // Extract width and height information from a BMP file.
    fn process_bmp(&mut self, data: &[u8]) {
        let width_dpi: f64 = 96.0;
        let height_dpi: f64 = 96.0;

        let width = unpack_u32_from_le_bytes(data, 18);
        let height = unpack_u32_from_le_bytes(data, 22);

        self.width = f64::from(width);
        self.height = f64::from(height);
        self.width_dpi = width_dpi;
        self.height_dpi = height_dpi;
        self.image_type = XlsxImageType::Bmp;
    }

    // Extract width and height information from a GIF file.
    fn process_gif(&mut self, data: &[u8]) {
        let width = u32::from(unpack_u16_from_le_bytes(data, 6));
        let height = u32::from(unpack_u16_from_le_bytes(data, 8));

        self.width = f64::from(width);
        self.height = f64::from(height);
        self.width_dpi = 96.0;
        self.height_dpi = 96.0;
        self.image_type = XlsxImageType::Gif;
    }
}

// Trait for objects that have a component stored in the drawing.xml file.
impl DrawingObject for Image {
    fn x_offset(&self) -> u32 {
        self.x_offset
    }

    fn y_offset(&self) -> u32 {
        self.y_offset
    }

    fn width_scaled(&self) -> f64 {
        self.width * self.scale_width * 96.0 / self.width_dpi
    }

    fn height_scaled(&self) -> f64 {
        self.height * self.scale_height * 96.0 / self.height_dpi
    }

    fn object_movement(&self) -> ObjectMovement {
        self.object_movement
    }

    fn name(&self) -> String {
        self.name.clone()
    }

    fn alt_text(&self) -> String {
        self.alt_text.clone()
    }

    fn decorative(&self) -> bool {
        self.decorative
    }

    fn drawing_type(&self) -> DrawingType {
        self.drawing_type
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------

/// The `ObjectMovement` enum defines the movement of worksheet objects such as
/// images and charts.
///
/// This enum defines the way control a worksheet object such as [Image],
/// [`Chart`](crate::Chart), [`Note`](crate::Note), [`Shape`](crate::Shape) or
/// [`Button`](crate::Button) moves when the cells underneath it are moved,
/// resized or deleted. This equates to the following Excel options:
///
/// <img src="https://rustxlsxwriter.github.io/images/object_movement.png">
///
/// Used with [`Image::set_object_movement()`].
///
#[derive(Clone, Debug, PartialEq, Eq, Copy)]
pub enum ObjectMovement {
    /// Move and size the worksheet object with the cells. Default for charts.
    MoveAndSizeWithCells,

    /// Move but don't size the worksheet object with the cells. Default for
    /// images.
    MoveButDontSizeWithCells,

    /// Don't move or size the worksheet object with the cells.
    DontMoveOrSizeWithCells,

    /// Same as `MoveAndSizeWithCells` except hidden cells are applied after the
    /// object is inserted. This allows the insertion of objects into hidden
    /// rows or columns.
    MoveAndSizeWithCellsAfter,
}

/// The `HeaderImagePosition` enum defines the image position in a header or footer.
///
/// Used with the
/// [`Worksheet::set_header_image()`](crate::Worksheet::set_header_image) and
/// [`Worksheet::set_footer_image()`](crate::Worksheet::set_footer_image)
/// methods.
///
#[derive(Clone, Debug)]
pub enum HeaderImagePosition {
    /// The image is positioned in the left section of the header/footer.
    Left,

    /// The image is positioned in the center section of the header/footer.
    Center,

    /// The image is positioned in the right section of the header/footer.
    Right,
}

#[derive(Clone, Debug)]
pub(crate) enum XlsxImageType {
    Unknown,
    Png,
    Jpg,
    Gif,
    Bmp,
}

impl XlsxImageType {
    pub(crate) fn extension(&self) -> String {
        match self {
            XlsxImageType::Unknown => "unknown".to_string(),
            XlsxImageType::Png => "png".to_string(),
            XlsxImageType::Jpg => "jpeg".to_string(),
            XlsxImageType::Gif => "gif".to_string(),
            XlsxImageType::Bmp => "bmp".to_string(),
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
