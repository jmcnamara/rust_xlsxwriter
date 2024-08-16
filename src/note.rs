// note - A module to represent Excel cell notes.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::drawing::{DrawingObject, DrawingType};
use crate::vml::VmlInfo;
use crate::{ColNum, Color, Format, ObjectMovement, RowNum, COL_MAX, ROW_MAX};

#[derive(Clone)]
/// The `Note` struct represents an worksheet note object.
///
/// A Note is a post-it style message that is revealed when the user mouses over
/// a worksheet cell. The presence of a Note is indicated by a small red
/// triangle in the upper right-hand corner of the cell.
///
/// <img src="https://rustxlsxwriter.github.io/images/app_notes.png">
///
/// The above file was created using the following code:
///
/// ```
/// # // This code is available in examples/app_notes.rs
/// #
/// use rust_xlsxwriter::{Note, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Widen the first column for clarity.
///     worksheet.set_column_width(0, 16)?;
///
///     // Write some data.
///     let party_items = [
///         "Invitations",
///         "Doors",
///         "Flowers",
///         "Champagne",
///         "Menu",
///         "Peter",
///     ];
///     worksheet.write_column(0, 0, party_items)?;
///
///     // Create a new worksheet Note.
///     let note = Note::new("I will get the flowers myself").set_author("Clarissa Dalloway");
///
///     // Add the note to a cell.
///     worksheet.insert_note(2, 0, &note)?;
///
///     // Save the file to disk.
///     workbook.save("notes.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Notes are used in conjunction with the
/// [`Worksheet::insert_note()`](crate::Worksheet::insert_note) method.
///
/// In versions of Excel prior to Office 365 Notes were referred to as
/// "Comments". The name Comment is now used for a newer style threaded comment
/// and Note is used for the older non threaded version. See the Microsoft docs
/// on [The difference between threaded comments and notes].
///
/// [The difference between threaded comments and notes]:
///     https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3
///
/// Note, the newer Threaded Comments are unlikely to be added to
/// `rust_xlsxwriter` due to the fact that the feature relies on company
/// specific metadata to identify the comment author.
///
pub struct Note {
    height: f64,
    width: f64,
    row: Option<RowNum>,
    col: Option<ColNum>,
    x_offset: Option<u32>,
    y_offset: Option<u32>,

    pub(crate) author: Option<String>,
    pub(crate) author_id: usize,
    pub(crate) has_author_prefix: bool,
    pub(crate) cell_row: RowNum,
    pub(crate) cell_col: ColNum,
    pub(crate) text: String,
    pub(crate) alt_text: String,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) decorative: bool,
    pub(crate) is_visible: Option<bool>,
    pub(crate) format: Format,
}

impl Note {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Note object to represent an Excel cell note.
    ///
    /// The text of the Note is added in the constructor.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_new.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note");
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/note_new.png">
    ///
    pub fn new(text: impl Into<String>) -> Note {
        let format = Format::new()
            .set_background_color("#FFFFE1")
            .set_font_name("Tahoma")
            .set_font_size(8);

        Note {
            row: None,
            col: None,
            x_offset: None,
            y_offset: None,

            cell_row: 0,
            cell_col: 0,

            author: None,
            author_id: 0,
            has_author_prefix: true,
            width: 128.0,
            height: 74.0,
            text: text.into(),
            alt_text: String::new(),
            object_movement: ObjectMovement::DontMoveOrSizeWithCells,
            decorative: false,
            is_visible: None,
            format,
        }
    }

    /// Set the note author name.
    ///
    /// The author name appears in two places: at the start of the note text in
    /// bold and at the bottom of the worksheet in the status bar.
    ///
    /// If no name is specified the default name "Author" will be applied to the
    /// note.
    ///
    /// You can also set the default author name for all notes in a worksheet
    /// via the
    /// [`Worksheet::set_default_note_author()`](crate::Worksheet::set_default_note_author)
    /// method.
    ///
    /// # Parameters
    ///
    /// - `name`: The note author name. Must be less than or equal to the Excel
    ///   limit of 52 characters.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell.
    /// This example also sets the author name.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_set_author.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note").set_author("Rust");
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/note_set_author.png">
    ///
    pub fn set_author(mut self, name: impl Into<String>) -> Note {
        let author = name.into();
        if author.chars().count() > 52 {
            eprintln!("Author name is greater than Excel's limit of 52 characters.");
            return self;
        }

        self.author = Some(author);
        self
    }

    /// Prefix the note text with the author name.
    ///
    /// By default Excel, and `rust_xlsxwriter`, prefixes the author name to the
    /// note text (see the previous examples). If you prefer to have the note
    /// text without the author name you can use this option to turn it off.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell.
    /// This example turns off the author name in the note.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_add_author_prefix.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note")
    ///         .add_author_prefix(false)
    ///         .set_author("Rust"); // This is ignored in the Note.
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/note_add_author_prefix.png">
    ///
    pub fn add_author_prefix(mut self, enable: bool) -> Note {
        self.has_author_prefix = enable;
        self
    }

    /// Reset the text in the note.
    ///
    /// In general the text of the note is set in the the [`Note::new()`]
    /// constructor but if required you can use the `reset_text()` method to
    /// reset the text for a note. This allows a single `Note` instance to be
    /// used multiple times and avoids the small overhead of creating a new
    /// instance each time.
    ///
    /// # Parameters
    ///
    /// - `text`: The text for the note.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell.
    /// This example reuses the Note object and reset the test.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_reset_text.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let mut note = Note::new("Some text for the note");
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    ///
    ///     // Reuse the Note with different text.
    ///     note.reset_text("Some other text");
    ///
    ///     // Add the note to another worksheet cell.
    ///     worksheet.insert_note(4, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/note_reset_text.png">
    ///
    ///
    pub fn reset_text(&mut self, text: impl Into<String>) -> &mut Note {
        self.text = text.into();
        self
    }

    /// Set the width of the note in pixels.
    ///
    /// The default width of an Excel note is 128 pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The note width in pixels.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell. This
    /// example also changes the note dimensions.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_set_width.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note")
    ///         .set_width(200)
    ///         .set_height(200);
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/note_set_width.png">
    ///
    pub fn set_width(mut self, width: u32) -> Note {
        if width == 0 {
            return self;
        }

        self.width = f64::from(width);
        self
    }

    /// Set the height of the note in pixels.
    ///
    /// The default height of an Excel note is 74 pixels. See the example above.
    ///
    /// # Parameters
    ///
    /// - `height`: The note height in pixels.
    ///
    pub fn set_height(mut self, height: u32) -> Note {
        if height == 0 {
            return self;
        }
        self.height = f64::from(height);
        self
    }

    /// Make the note visible when the file loads.
    ///
    /// By default Excel hides cell notes until the user mouses over the parent
    /// cell. However, if required you can make the note visible without
    /// requiring an interaction from the user.
    ///
    /// You can also make all notes in a worksheet visible via the
    /// [`Worksheet::show_all_notes()`](crate::Worksheet::show_all_notes)
    /// method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell.
    /// This example makes the note visible by default.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_set_visible.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note").set_visible(true);
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/note_set_visible.png">
    ///
    pub fn set_visible(mut self, enable: bool) -> Note {
        self.is_visible = Some(enable);
        self
    }

    /// Set the background color for the note.
    ///
    /// The default background color for a Note is `#FFFFE1`. If required this
    /// method can be used to set it to a different RGB color.
    ///
    /// # Parameters
    ///
    /// - `color`: The background color property defined by a
    ///   [`Color`](crate::Color) enum value or a type that can convert [`Into`]
    ///   a [`Color`]. Only the `Color::Name` and `Color::RGB()` variants are
    ///   supported. Theme style colors aren't support by Excel for Notes.
    ///
    /// # Examples
    ///
    /// The following example demonstrates adding a note to a worksheet cell.
    /// This example also sets the background color.
    ///
    /// ```
    /// # // This code is available in examples/doc_note_set_background_color.rs
    /// #
    /// # use rust_xlsxwriter::{Note, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a new note.
    ///     let note = Note::new("Some text for the note").set_background_color("#80ff80");
    ///
    ///     // Add the note to a worksheet cell.
    ///     worksheet.insert_note(2, 0, &note)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("notes.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/note_set_background_color.png">
    ///
    pub fn set_background_color(mut self, color: impl Into<Color>) -> Note {
        let color = color.into();
        if color.is_valid() {
            self.format.fill.background_color = color;
        }

        self
    }

    /// Set the font name for the note.
    ///
    /// Set the font for a cell note. Excel can only display fonts that are
    /// installed on the system that it is running on. Therefore it is generally
    /// best to use standard Excel fonts.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name for the note.
    ///
    pub fn set_font_name(mut self, font_name: impl Into<String>) -> Note {
        self.format.font.name = font_name.into();

        if self.format.font.name != "Calibri" {
            self.format.font.scheme = String::new();
        }

        self
    }

    /// Set the font size for the note.
    ///
    /// Set the font size of the cell format. The size is generally an integer
    /// value but Excel allows x.5 values (hence the property is a f64 or
    /// types that can convert [`Into`] a f64).
    ///
    /// # Parameters
    ///
    /// - `font_size`: The font size for the note.
    ///
    pub fn set_font_size<T>(mut self, font_size: T) -> Note
    where
        T: Into<f64>,
    {
        self.format.font.size = font_size.into().to_string();
        self
    }

    /// Set the font family for the note.
    ///
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    ///
    /// # Parameters
    ///
    /// - `font_family`: The font family for the note.
    ///
    #[doc(hidden)]
    pub fn set_font_family(mut self, font_family: u8) -> Note {
        self.format.font.family = font_family;
        self
    }

    /// Set the [`Format`] of the note.
    ///
    /// Set the font or background properties of a note using a [`Format`]
    /// object. Only the font name, size and background color are supported.
    ///
    /// This API is currently experimental and may go away in the future.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`Format`] property for the note.
    ///
    #[doc(hidden)]
    pub fn set_format(mut self, format: impl Into<Format>) -> Note {
        self.format = format.into();
        self
    }

    /// Set the alt text for the note to help accessibility.
    ///
    /// The alt text is used with screen readers to help people with visual
    /// disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the note.
    ///
    pub fn set_alt_text(mut self, alt_text: impl Into<String>) -> Note {
        let alt_text = alt_text.into();
        if alt_text.chars().count() > 255 {
            eprintln!("Alternative text is greater than Excel's limit of 255 characters.");
            return self;
        }

        self.alt_text = alt_text;
        self
    }

    /// Set the object movement options for a worksheet note.
    ///
    /// Set the option to define how an note will behave in Excel if the cells
    /// under the note are moved, deleted, or have their size changed. In
    /// Excel the options are:
    ///
    /// 1. Move and size with cells.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/object_movement.png">
    ///
    /// These values are defined in the [`ObjectMovement`] enum.
    ///
    /// The [`ObjectMovement`] enum also provides an additional option to "Move
    /// and size with cells - after the note is inserted" to allow notes to
    /// be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal note position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// - `option`: An note/object positioning behavior defined by the
    ///   [`ObjectMovement`] enum.
    pub fn set_object_movement(mut self, option: ObjectMovement) -> Note {
        self.object_movement = option;
        self
    }

    // Notes are stored in a vmlDrawing file. We create a struct to store the
    // required image information in that format.
    pub(crate) fn vml_info(&self) -> VmlInfo {
        VmlInfo {
            width: self.width,
            height: self.height,
            text: self.text.clone(),
            alt_text: self.alt_text.clone(),
            is_visible: self.is_visible.unwrap_or(false),
            fill_color: self.format.fill.background_color.vml_rgb_hex_value(),
            ..Default::default()
        }
    }

    // Notes have a stand row offset relative to the parent cell except at the
    // top and bottom of the worksheet.
    pub(crate) fn row(&self) -> RowNum {
        match self.row {
            Some(row) => row,
            None => {
                if self.cell_row == 0 {
                    0
                } else if self.cell_row == ROW_MAX - 3 {
                    ROW_MAX - 7
                } else if self.cell_row == ROW_MAX - 2 {
                    ROW_MAX - 6
                } else if self.cell_row == ROW_MAX - 1 {
                    ROW_MAX - 5
                } else {
                    self.cell_row - 1
                }
            }
        }
    }

    // Notes have a stand column offset relative to the parent cell except at
    // the right side of the worksheet.
    pub(crate) fn col(&self) -> ColNum {
        match self.col {
            Some(col) => col,
            None => {
                if self.cell_col == COL_MAX - 3 {
                    COL_MAX - 6
                } else if self.cell_col == COL_MAX - 2 {
                    COL_MAX - 5
                } else if self.cell_col == COL_MAX - 1 {
                    COL_MAX - 4
                } else {
                    self.cell_col + 1
                }
            }
        }
    }
}

// Trait for objects that have a component stored in the drawing.xml file.
impl DrawingObject for Note {
    #[allow(clippy::if_same_then_else)]
    fn x_offset(&self) -> u32 {
        match self.x_offset {
            Some(offset) => offset,
            None => {
                if self.cell_col == COL_MAX - 3 {
                    49
                } else if self.cell_col == COL_MAX - 2 {
                    49
                } else if self.cell_col == COL_MAX - 1 {
                    49
                } else {
                    15
                }
            }
        }
    }

    #[allow(clippy::if_same_then_else)]
    fn y_offset(&self) -> u32 {
        match self.y_offset {
            Some(offset) => offset,
            None => {
                if self.cell_row == 0 {
                    2
                } else if self.cell_row == ROW_MAX - 3 {
                    16
                } else if self.cell_row == ROW_MAX - 2 {
                    16
                } else if self.cell_row == ROW_MAX - 1 {
                    14
                } else {
                    10
                }
            }
        }
    }

    fn width_scaled(&self) -> f64 {
        self.width
    }

    fn height_scaled(&self) -> f64 {
        self.height
    }

    fn object_movement(&self) -> ObjectMovement {
        self.object_movement
    }

    fn name(&self) -> String {
        self.text.clone()
    }

    fn alt_text(&self) -> String {
        self.alt_text.clone()
    }

    fn decorative(&self) -> bool {
        self.decorative
    }

    fn drawing_type(&self) -> DrawingType {
        DrawingType::Vml
    }
}
