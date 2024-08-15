// protection - A module for representing worksheet protection options.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

/// The `ProtectionOptions` struct is use to set protected elements in a worksheet.
///
/// You can specify which worksheet elements protection should be on or off via
/// the `ProtectionOptions` members. The corresponding Excel options with
/// their default states are shown below:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options1.png">
///
/// # Examples
///
/// The following example demonstrates setting the worksheet properties to be
/// protected in a protected worksheet. In this case we protect the overall
/// worksheet but allow columns and rows to be inserted.
///
/// ```
/// # // This code is available in examples/doc_worksheet_protect_with_options.rs
/// #
/// use rust_xlsxwriter::{ProtectionOptions, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Set some of the options and use the defaults for everything else.
///     let options = ProtectionOptions {
///         insert_columns: true,
///         insert_rows: true,
///         ..ProtectionOptions::default()
///     };
///
///     // Set the protection options.
///     worksheet.protect_with_options(&options);
///
///     worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;
///
///     workbook.save("worksheet.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Excel dialog for the output file, compare this with the default image above:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options2.png">
///
///
///
#[derive(Clone)]
pub struct ProtectionOptions {
    /// When `true` (the default) the user can select locked cells in a
    /// protected worksheet.
    pub select_locked_cells: bool,

    /// When `true` (the default) the user can select unlocked cells in a
    /// protected worksheet.
    pub select_unlocked_cells: bool,

    /// When `false` (the default) the user cannot format cells in a protected
    /// worksheet.
    pub format_cells: bool,

    /// When `false` (the default) the user cannot format cells in a protected
    /// worksheet.
    pub format_columns: bool,

    /// When `false` (the default) the user cannot format rows in a protected
    /// worksheet.
    pub format_rows: bool,

    /// When `false` (the default) the user cannot insert new columns in a
    /// protected worksheet.
    pub insert_columns: bool,

    /// When `false` (the default) the user cannot insert new rows in a
    /// protected worksheet.
    pub insert_rows: bool,

    /// When `false` (the default) the user cannot insert hyperlinks/urls in a
    /// protected worksheet.
    pub insert_links: bool,

    /// When `false` (the default) the user cannot delete columns in a protected
    /// worksheet.
    pub delete_columns: bool,

    /// When `false` (the default) the user cannot delete rows in a protected
    /// worksheet.
    pub delete_rows: bool,

    /// When `false` (the default) the user cannot sort data in a protected
    /// worksheet.
    pub sort: bool,

    /// When `false` (the default) the user cannot use autofilters in a
    /// protected worksheet.
    pub use_autofilter: bool,

    /// When `false` (the default) the user cannot use pivot tables or pivot
    /// charts in a protected worksheet.
    pub use_pivot_tables: bool,

    /// When `false` (the default) the user cannot edit scenarios in a protected
    /// worksheet.
    pub edit_scenarios: bool,

    /// When `false` (the default) the user cannot edit objects such as images,
    /// charts or textboxes in a protected worksheet.
    pub edit_objects: bool,
}

impl Default for ProtectionOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl ProtectionOptions {
    /// Create a new [`ProtectionOptions`] object to use with the
    /// [`Worksheet::protect_with_options()`](crate::Worksheet::protect_with_options) method.
    ///
    pub fn new() -> ProtectionOptions {
        ProtectionOptions {
            select_locked_cells: true,
            select_unlocked_cells: true,
            format_cells: false,
            format_columns: false,
            format_rows: false,
            insert_columns: false,
            insert_rows: false,
            insert_links: false,

            delete_columns: false,
            delete_rows: false,
            sort: false,
            use_autofilter: false,
            use_pivot_tables: false,
            edit_scenarios: false,
            edit_objects: false,
        }
    }
}
