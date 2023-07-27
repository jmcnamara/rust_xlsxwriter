// table - A module for creating the Excel Table.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::{collections::HashSet, fmt};

use crate::{
    utility::ToXmlBoolean, xmlwriter::XMLWriter, CellRange, Format, Formula, RowNum, XlsxError,
};

/// The `Table` struct represents a worksheet Table.
///
/// Tables in Excel are a way of grouping a range of cells into a single entity
/// that has common formatting or that can be referenced from formulas. Tables
/// can have column headers, autofilters, total rows, column formulas and
/// different formatting styles.
///
/// The image below shows a default table in Excel with the default properties
/// shown in the ribbon bar.
///
/// <img src="https://rustxlsxwriter.github.io/images/table_intro.png">
///
/// A table is added to a worksheet via the
/// [`worksheet.add_table()`](crate::Worksheet::add_table) method. The headers
/// and total row of a table should be configured via a `Table` struct but the
/// table data can be added via standard
/// [`worksheet.write()`](crate::Worksheet::write) methods:
///
/// ```
/// # // This code is available in examples/doc_table_set_columns.rs
/// #
/// use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Some sample data for the table.
///     let items = ["Apples", "Pears", "Bananas", "Oranges"];
///     let data = [
///         [10000, 5000, 8000, 6000],
///         [2000, 3000, 4000, 5000],
///         [6000, 6000, 6500, 6000],
///         [500, 300, 200, 700],
///     ];
///
///     // Write the table data.
///     worksheet.write_column(3, 1, items)?;
///     worksheet.write_row_matrix(3, 2, data)?;
///
///     // Set the columns widths for clarity.
///     for col_num in 1..=6u16 {
///         worksheet.set_column_width(col_num, 12)?;
///     }
///
///     // Create a new table and configure it.
///     let mut table = Table::new();
///
///     let columns = vec![
///         TableColumn::new()
///             .set_header("Product")
///             .set_total_label("Totals"),
///         TableColumn::new()
///             .set_header("Quarter 1")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Quarter 2")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Quarter 3")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Quarter 4")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Year")
///             .set_total_function(TableFunction::Sum)
///             .set_formula("SUM(Table1[@[Quarter 1]:[Quarter 4]])"),
///     ];
///
///     table.set_columns(&columns);
///     table.set_total_row(true);
///
///     // Add the table to the worksheet.
///     worksheet.add_table(2, 1, 7, 6, &table)?;
///
///     // Save the file to disk.
///     workbook.save("tables.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/table_set_columns.png">
///
/// For more information on tables see the Microsoft documentation on [Overview
/// of Excel tables].
///
/// [Overview of Excel tables]:
///     https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
///
#[derive(Clone)]
pub struct Table {
    pub(crate) writer: XMLWriter,

    pub(crate) columns: Vec<TableColumn>,

    pub(crate) index: u32,
    pub(crate) name: String,
    pub(crate) style: TableStyle,

    pub(crate) cell_range: CellRange,

    pub(crate) show_header_row: bool,
    pub(crate) show_total_row: bool,
    pub(crate) show_first_column: bool,
    pub(crate) show_last_column: bool,
    pub(crate) show_banded_rows: bool,
    pub(crate) show_banded_columns: bool,
    pub(crate) show_autofilter: bool,
}

impl Table {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Table struct instance.
    ///
    /// Create a table that can be added to a data range of a worksheet. The
    /// headers, totals, formulas and other properties can be set via the
    /// `Table::*` methods shown below. The data should be added to the table
    /// region using the standard
    /// [`Worksheet::write()`](crate::Worksheet::write) methods.
    ///
    /// # Examples
    ///
    /// Example of creating a new table and adding it to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_header_row2.rs
    /// #
    /// use rust_xlsxwriter::{Table, Workbook, XlsxError};
    ///
    /// fn main() -> Result<(), XlsxError> {
    ///     // Create a new Excel file object.
    ///     let mut workbook = Workbook::new();
    ///
    ///     // Add a worksheet to the workbook.
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Some sample data for the table.
    ///     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    ///     let data = [
    ///         [10000, 5000, 8000, 6000],
    ///         [2000, 3000, 4000, 5000],
    ///         [6000, 6000, 6500, 6000],
    ///         [500, 300, 200, 700],
    ///     ];
    ///
    ///     // Write the table data.
    ///     worksheet.write_column(3, 1, items)?;
    ///     worksheet.write_row_matrix(3, 2, data)?;
    ///
    ///     // Set the columns widths for clarity.
    ///     for col_num in 1..=6u16 {
    ///         worksheet.set_column_width(col_num, 12)?;
    ///     }
    ///
    ///     // Create a new table.
    ///     let table = Table::new();
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("tables.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row2.png">
    ///
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> Table {
        let writer = XMLWriter::new();

        Table {
            writer,
            columns: vec![],
            index: 0,
            name: String::new(),
            style: TableStyle::Medium9,
            cell_range: CellRange::default(),
            show_first_column: false,
            show_last_column: false,
            show_banded_rows: true,
            show_banded_columns: false,
            show_autofilter: true,
            show_header_row: true,
            show_total_row: false,
        }
    }

    /// Turn on/off the header row as a table.
    ///
    /// Turn on or off the header row in the table.  The header row displays the
    /// column names and, unless it is turned off, an autofilter. It is on by
    /// default.
    ///
    /// The header row will display default captions such as `Column 1`, `Column
    /// 2`, etc. These captions can be overridden using the
    /// [`set_columns()`](Table::set_columns) method, see the examples below.
    /// They shouldn't be written, or overwritten using standard
    /// [`Worksheet::write()`](crate::Worksheet::write) methods since that will
    /// cause a warning when the file is loaded in Excel.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// Example of adding a worksheet table with a default header.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_header_row2.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table.
    ///     let table = Table::new();
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row2.png">
    ///
    ///
    /// Example of turning off the default header on a worksheet table. Note,
    /// that the table range has been adjusted in relation to the previous
    /// example to account for the missing header.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_header_row.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(2, 1, items)?;
    /// #     worksheet.write_row_matrix(2, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the header.
    ///     let mut table = Table::new();
    ///     table.set_header_row(false);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 5, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row.png">
    ///
    /// Example of adding a worksheet table with a user defined header captions.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_header_row3.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the column headers.
    ///     let mut table = Table::new();
    ///
    ///     // Set the captions for the header row.
    ///     let columns = vec![
    ///         TableColumn::new().set_header("Product"),
    ///         TableColumn::new().set_header("Quarter 1"),
    ///         TableColumn::new().set_header("Quarter 2"),
    ///         TableColumn::new().set_header("Quarter 3"),
    ///         TableColumn::new().set_header("Quarter 4"),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row3.png">
    ///
    pub fn set_header_row(&mut self, enable: bool) -> &mut Table {
        self.show_header_row = enable;
        self
    }

    /// Turn on a totals row for a table.
    ///
    /// The `set_total_row()` method can be used to turn on the total row in the
    /// last row of a table. The total row is distinguished from the other rows
    /// by a different formatting and with dropdown `SUBTOTAL()` functions.
    ///
    /// Note, you will need to use [`TableColumn`] methods to populate this row.
    /// Overwriting the total row cells with `worksheet.write()` calls will
    /// cause Excel to warn that the table is corrupt when loading the file.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    ///
    /// # Examples
    ///
    /// Example of turning on the "totals" row at the bottom of a worksheet
    /// table. Note, this just turns on the total run it doesn't add captions or
    /// subtotal functions. See the next example below.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_total_row.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the total row.
    ///     let mut table = Table::new();
    ///     table.set_total_row(true);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 7, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_total_row.png">
    ///
    /// Example of turning on the "totals" row at the bottom of a worksheet
    /// table with captions and subtotal functions.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_total_row2.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the total row.
    ///     let mut table = Table::new();
    ///     table.set_total_row(true);
    ///
    ///     // Set the caption and subtotal in the total row.
    ///     let columns = vec![
    ///         TableColumn::new().set_total_label("Totals"),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 7, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_total_row2.png">
    ///
    pub fn set_total_row(&mut self, enable: bool) -> &mut Table {
        self.show_total_row = enable;
        self
    }

    /// Turn on/off banded for a table.
    ///
    /// By default Excel uses "banded" rows of alternating colors in a table to
    /// distinguish each data row, like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row2.png">
    ///
    /// If you prefer not to have this type of formatting you can turn it off,
    /// see the example below.
    ///
    /// Note, you can also select a table style without banded rows using the
    /// [`table.set_style`](Table::set_style) method.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    /// # Examples
    ///
    /// Example of turning off the banded rows property in a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_banded_rows.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the banded rows.
    ///     let mut table = Table::new();
    ///     table.set_banded_rows(false);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_banded_rows.png">
    ///
    pub fn set_banded_rows(&mut self, enable: bool) -> &mut Table {
        self.show_banded_rows = enable;
        self
    }

    /// Turn on/off banded columns for a table.
    ///
    /// By default Excel uses the same format color for each data column in a
    /// table but alternates the color of rows. If you wish you can set "banded"
    /// columns of alternating colors in a table to distinguish each data column.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// Example of turning on the banded columns property in a worksheet table. These
    /// are normally off by default,
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_banded_columns.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the banded columns.
    ///     let mut table = Table::new();
    ///     table.set_banded_columns(true);
    ///
    ///     // Turn off banded rows for clarity.
    ///     table.set_banded_rows(false);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/table_set_banded_columns.png">
    ///
    pub fn set_banded_columns(&mut self, enable: bool) -> &mut Table {
        self.show_banded_columns = enable;
        self
    }

    /// Turn on/off the first column highlighting for a table.
    ///
    /// The first column of a worksheet table is often used for a list of items
    /// whereas the other columns are more commonly used for data. In these
    /// cases it is sometimes desirable to highlight the first column differently.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// Example of turning on the first column highlighting property in a
    /// worksheet table. This is normally off by default,
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_first_column.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the first column highlighting.
    ///     let mut table = Table::new();
    ///     table.set_first_column(true);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_first_column.png">
    ///
    pub fn set_first_column(&mut self, enable: bool) -> &mut Table {
        self.show_first_column = enable;
        self
    }

    /// Turn on/off the last column highlighting for a table.
    ///
    /// The last column of a worksheet table is often used for a `SUM()` or
    /// other formula operating on the  data in the other columns. In these
    /// cases it is sometimes required to highlight the last column differently.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// Example of turning on the last column highlighting property in a
    /// worksheet table. This is normally off by default,
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_last_column.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the last column highlighting.
    ///     let mut table = Table::new();
    ///     table.set_last_column(true);
    ///
    ///     // Add a structured reference formula to the last column and set the header
    ///     // caption. The last column in `add_table()` should be extended to account
    ///     // for this extra column.
    ///     let columns = vec![
    ///         TableColumn::default(),
    ///         TableColumn::default(),
    ///         TableColumn::default(),
    ///         TableColumn::default(),
    ///         TableColumn::default(),
    ///         TableColumn::new()
    ///             .set_header("Totals")
    ///             .set_formula("SUM(Table1[@[Column2]:[Column5]])"),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 6, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_last_column.png">
    ///
    pub fn set_last_column(&mut self, enable: bool) -> &mut Table {
        self.show_last_column = enable;
        self
    }

    /// Turn on/off the autofilter for a table.
    ///
    /// By default Excel adds an autofilter to the header of a table. This
    /// method can be used to turn it off if necessary.
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// Example of turning off the autofilter in a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_autofilter.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the autofilter.
    ///     let mut table = Table::new();
    ///     table.set_autofilter(false);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/table_set_autofilter.png">
    ///
    pub fn set_autofilter(&mut self, enable: bool) -> &mut Table {
        self.show_autofilter = enable;
        self
    }

    /// Set properties for the columns in a table.
    ///
    /// Set the properties for columns in a worksheet table via an array of
    /// [`TableColumn`] structs. This can be used to set the following
    /// properties of a table column:
    ///
    /// - The header caption.
    /// - The total row caption.
    /// - The total row subtotal function.
    /// - A formula for the column.
    ///
    ///
    /// # Parameters
    ///
    /// * `columns` - An array reference of [`TableColumn`] structs. Use
    ///   `TableColumn::default()` to get default values.
    ///
    ///
    /// # Examples
    ///
    /// Example of creating a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_columns.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure it.
    ///     let mut table = Table::new();
    ///
    ///     let columns = vec![
    ///         TableColumn::new()
    ///             .set_header("Product")
    ///             .set_total_label("Totals"),
    ///         TableColumn::new()
    ///             .set_header("Quarter 1")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Quarter 2")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Quarter 3")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Quarter 4")
    ///             .set_total_function(TableFunction::Sum),
    ///         TableColumn::new()
    ///             .set_header("Year")
    ///             .set_total_function(TableFunction::Sum)
    ///             .set_formula("SUM(Table1[@[Quarter 1]:[Quarter 4]])"),
    ///     ];
    ///
    ///     table.set_columns(&columns);
    ///     table.set_total_row(true);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 7, 6, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_columns.png">
    ///
    pub fn set_columns(&mut self, columns: &[TableColumn]) -> &mut Table {
        self.columns = columns.to_vec();
        self
    }

    /// Set the name for a table.
    ///
    /// The name of a worksheet table in Excel is similar to a defined name
    /// representing a data region and it can be used in structured reference
    /// formulas.
    ///
    /// By default Excel, and `rust_xlsxwriter` uses a `Table1` .. `TableN`
    /// naming for tables in a workbook. If required you can set a user defined
    /// name. However, you need to ensure that this name is unique across the
    /// workbook, otherwise you will get a warning when you load the file in
    /// Excel.
    ///
    /// # Parameters
    ///
    /// * `name` - The name of the table. It must be unique across the workbook.
    ///
    /// # Examples
    ///
    /// Example of setting the name of a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Table, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and set the name.
    ///     let mut table = Table::new();
    ///     table.set_name("ProduceSales");
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/table_set_name.png">
    ///
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Table {
        self.name = name.into();
        self
    }

    /// Set the style for a table.
    ///
    /// Excel supports 61 different styles for tables divided into Light, Medium
    /// and Dark categories.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/table_styles.png">
    ///
    /// You can set one of these styles using a [`TableStyle`] enum value. The
    /// default table style in Excel is equivalent to
    /// [`TableStyle::Medium9`](TableStyle::Medium9).
    ///
    /// # Parameters
    ///
    /// * `style` - a [`TableStyle`] enum value.
    ///
    /// # Examples
    ///
    /// Example of setting the style of a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_style.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableStyle, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and set the style.
    ///     let mut table = Table::new();
    ///     table.set_style(TableStyle::Medium10);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/table_set_style.png">
    ///
    pub fn set_style(&mut self, style: TableStyle) -> &mut Table {
        self.style = style;
        self
    }

    // Truncate or extend (with defaults) the table columns.
    pub(crate) fn initialize_columns(
        &mut self,
        default_headers: &[String],
    ) -> Result<(), XlsxError> {
        let mut seen_column_names = HashSet::new();
        let num_columns = self.cell_range.last_col - self.cell_range.first_col + 1;

        self.columns
            .resize_with(num_columns as usize, TableColumn::default);

        // Set the column header names,
        for (index, column) in self.columns.iter_mut().enumerate() {
            if column.name.is_empty() {
                column.name = default_headers[index].clone();
            }

            if seen_column_names.contains(&column.name.to_lowercase()) {
                return Err(XlsxError::TableError(format!(
                    "Column name '{}' already exists in Table at {}",
                    column.name,
                    self.cell_range.to_error_string()
                )));
            }

            seen_column_names.insert(column.name.to_lowercase().clone());
        }

        Ok(())
    }

    // Get the first row that can be used to write data.
    pub(crate) fn first_data_row(&self) -> RowNum {
        if self.show_header_row {
            self.cell_range.first_row + 1
        } else {
            self.cell_range.first_row
        }
    }

    // Get the last row that can be used to write data.
    pub(crate) fn last_data_row(&self) -> RowNum {
        if self.show_total_row {
            self.cell_range.last_row - 1
        } else {
            self.cell_range.last_row
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the table element.
        self.write_table();

        if self.show_autofilter && self.show_header_row {
            // Write the autoFilter element.
            self.write_auto_filter();
        }

        // Write the tableColumns element.
        self.write_columns();

        // Write the tableStyleInfo element.
        self.write_table_style_info();

        // Close the table tag.
        self.writer.xml_end_tag("table");
    }

    // Write the <table> element.
    fn write_table(&mut self) {
        let schema = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let range = self.cell_range.to_range_string();
        let name = if self.name.is_empty() {
            format!("Table{}", self.index)
        } else {
            self.name.clone()
        };

        let mut attributes = vec![
            ("xmlns", schema),
            ("id", self.index.to_string()),
            ("name", name.clone()),
            ("displayName", name),
            ("ref", range),
        ];

        if !self.show_header_row {
            attributes.push(("headerRowCount", "0".to_string()));
        }

        if self.show_total_row {
            attributes.push(("totalsRowCount", "1".to_string()));
        } else {
            attributes.push(("totalsRowShown", "0".to_string()));
        }

        self.writer.xml_start_tag("table", &attributes);
    }

    // Write the <autoFilter> element.
    fn write_auto_filter(&mut self) {
        let mut autofilter_range = self.cell_range.clone();

        if self.show_total_row {
            autofilter_range.last_row -= 1;
        }

        let attributes = vec![("ref", autofilter_range.to_range_string())];

        self.writer.xml_empty_tag("autoFilter", &attributes);
    }

    // Write the <tableColumns> element.
    fn write_columns(&mut self) {
        let attributes = vec![("count", self.columns.len().to_string())];

        self.writer.xml_start_tag("tableColumns", &attributes);

        for (index, column) in self.columns.clone().iter().enumerate() {
            // Write the tableColumn element.
            self.write_column(index + 1, column);
        }

        self.writer.xml_end_tag("tableColumns");
    }

    // Write the <tableColumn> element.
    fn write_column(&mut self, index: usize, column: &TableColumn) {
        let mut attributes = vec![("id", index.to_string()), ("name", column.name.clone())];

        if !column.total_label.is_empty() {
            attributes.push(("totalsRowLabel", column.total_label.clone()));
        } else if column.total_function != TableFunction::None {
            attributes.push(("totalsRowFunction", column.total_function.to_string()));
        }

        if let Some(format) = &column.format {
            attributes.push(("dataDxfId", format.dxf_index.to_string()));
        }

        if let Some(formula) = &column.formula {
            self.writer.xml_start_tag("tableColumn", &attributes);

            // Write the calculatedColumnFormula element.
            self.write_calculated_column_formula(&formula.expand_formula(true));

            self.writer.xml_end_tag("tableColumn");
        } else {
            self.writer.xml_empty_tag("tableColumn", &attributes);
        }
    }

    // Write the <calculatedColumnFormula> element.
    fn write_calculated_column_formula(&mut self, formula: &str) {
        self.writer
            .xml_data_element_only("calculatedColumnFormula", formula);
    }

    // Write the <tableStyleInfo> element.
    fn write_table_style_info(&mut self) {
        let mut attributes = vec![];

        if self.style != TableStyle::None {
            attributes.push(("name", self.style.to_string()));
        }

        attributes.push(("showFirstColumn", self.show_first_column.to_xml_bool()));
        attributes.push(("showLastColumn", self.show_last_column.to_xml_bool()));
        attributes.push(("showRowStripes", self.show_banded_rows.to_xml_bool()));
        attributes.push(("showColumnStripes", self.show_banded_columns.to_xml_bool()));

        self.writer.xml_empty_tag("tableStyleInfo", &attributes);
    }
}

#[derive(Clone)]
/// The `TableColumn` struct represents a table column.
///
/// The `TableColumn` struct is used to set the properties for columns in a
/// worksheet [`Table`]. This can be used to set the following properties of a
/// table column:
///
/// - The header caption.
/// - The total row caption.
/// - The total row subtotal function.
/// - A formula for the column.
///
/// This struct is used in conjunction with the
/// [`table.set_columns()`](Table::set_columns) method.
///
/// # Examples
///
/// Example of creating a worksheet table.
///
/// ```
/// # // This code is available in examples/doc_table_set_columns.rs
/// #
/// use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Some sample data for the table.
///     let items = ["Apples", "Pears", "Bananas", "Oranges"];
///     let data = [
///         [10000, 5000, 8000, 6000],
///         [2000, 3000, 4000, 5000],
///         [6000, 6000, 6500, 6000],
///         [500, 300, 200, 700],
///     ];
///
///     // Write the table data.
///     worksheet.write_column(3, 1, items)?;
///     worksheet.write_row_matrix(3, 2, data)?;
///
///     // Set the columns widths for clarity.
///     for col_num in 1..=6u16 {
///         worksheet.set_column_width(col_num, 12)?;
///     }
///
///     // Create a new table and configure it.
///     let mut table = Table::new();
///
///     let columns = vec![
///         TableColumn::new()
///             .set_header("Product")
///             .set_total_label("Totals"),
///         TableColumn::new()
///             .set_header("Quarter 1")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Quarter 2")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Quarter 3")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Quarter 4")
///             .set_total_function(TableFunction::Sum),
///         TableColumn::new()
///             .set_header("Year")
///             .set_total_function(TableFunction::Sum)
///             .set_formula("SUM(Table1[@[Quarter 1]:[Quarter 4]])"),
///     ];
///
///     table.set_columns(&columns);
///     table.set_total_row(true);
///
///     // Add the table to the worksheet.
///     worksheet.add_table(2, 1, 7, 6, &table)?;
///
///     // Save the file to disk.
///     workbook.save("tables.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/table_set_columns.png">
///
///
pub struct TableColumn {
    pub(crate) name: String,
    pub(crate) total_function: TableFunction,
    pub(crate) total_label: String,
    pub(crate) formula: Option<Formula>,
    pub(crate) format: Option<Format>,
    pub(crate) header_format: Option<Format>,
}

impl TableColumn {
    /// Create a new `TableColumn` to configure a Table column.
    ///
    pub fn new() -> TableColumn {
        TableColumn {
            name: String::new(),
            total_function: TableFunction::None,
            total_label: String::new(),
            formula: None,
            format: None,
            header_format: None,
        }
    }

    /// Set the header caption for a table column.
    ///
    /// Excel uses default captions such as `Column 1`, `Column 2`, etc., for
    /// the headers on a worksheet table. These can be set to a user defined
    /// value using the `set_header()` method.
    ///
    /// The column header names in a table must be different from each other.
    /// Non-unique names will raise a validation error when using
    /// [`worksheet.add_table()`](crate::Worksheet::add_table).
    ///
    /// # Parameters
    ///
    /// * `caption` - The caption/name of the column header. It must be unique
    ///   for the table.
    ///
    /// # Examples
    ///
    /// Example of adding a worksheet table with a user defined header captions.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_header_row3.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the column headers.
    ///     let mut table = Table::new();
    ///
    ///     // Set the captions for the header row.
    ///     let columns = vec![
    ///         TableColumn::new().set_header("Product"),
    ///         TableColumn::new().set_header("Quarter 1"),
    ///         TableColumn::new().set_header("Quarter 2"),
    ///         TableColumn::new().set_header("Quarter 3"),
    ///         TableColumn::new().set_header("Quarter 4"),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row3.png">
    ///
    pub fn set_header(mut self, caption: impl Into<String>) -> TableColumn {
        self.name = caption.into();
        self
    }

    /// Set the total function for the total row of a table column.
    ///
    /// Set the `SUBTOTAL()` function for the "totals" row of a table column.
    ///
    /// The standard Excel subtotal functions are available via the
    /// [`TableFunction`] enum values. The Excel functions are:
    ///
    /// - Average
    /// - Count
    /// - Count Numbers
    /// - Maximum
    /// - Minimum
    /// - Sum
    /// - Standard Deviation
    /// - Variance
    ///
    /// Excel also supports custom functions. These aren't currently supported
    /// by `rust_xlsxwriter` but will be added in an upcoming release.
    ///
    /// Note, overwriting the total row cells with `worksheet.write()` calls
    /// will cause Excel to warn that the table is corrupt when loading the
    /// file.
    ///
    /// # Parameters
    ///
    /// * `function` - A [`TableFunction`] enum value equivalent to one of the
    ///   available Excel `SUBTOTAL()` options.
    ///
    /// # Examples
    ///
    /// Example of turning on the "totals" row at the bottom of a worksheet
    /// table with captions and subtotal functions.
    ///
    /// ```
    /// # // This code is available in examples/doc_table_set_total_row2.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the total row.
    ///     let mut table = Table::new();
    ///     table.set_total_row(true);
    ///
    ///     // Set the caption and subtotal in the total row.
    ///     let columns = vec![
    ///         TableColumn::new().set_total_label("Totals"),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///         TableColumn::new().set_total_function(TableFunction::Sum),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 7, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/table_set_total_row2.png">
    ///
    pub fn set_total_function(mut self, function: TableFunction) -> TableColumn {
        self.total_function = function;
        self
    }

    /// Set a label for the total row of a table column.
    ///
    /// It is possible to set a label for the totals row of a column instead of
    /// a subtotal function. This is most often used to set a caption like
    /// "Totals", as in the example above.
    ///
    /// Note, overwriting the total row cells with `worksheet.write()` calls
    /// will cause Excel to warn that the table is corrupt when loading the
    /// file.
    ///
    /// # Parameters
    ///
    /// * `label` - The label/caption of the total row of the column.
    ///
    pub fn set_total_label(mut self, label: impl Into<String>) -> TableColumn {
        self.total_label = label.into();
        self
    }

    /// Set the formula for a table column.
    ///
    /// It is a common use case to add a summation column as the last column in a
    /// table. These are constructed with a special class of Excel formulas
    /// called [Structured References] which can refer to an entire table or
    /// rows or columns of data within the table. For example to sum the data
    /// for several columns in a single row might you might use a formula like
    /// this: `SUM(Table1[@[Quarter 1]:[Quarter 4]])`.
    ///
    /// [Structured References]:
    ///     https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    ///
    /// # Parameters
    ///
    /// * `formula` - The formula to be applied to the column as a string or
    ///   [`Formula`].
    ///
    /// # Examples
    ///
    /// Example of adding a formula to a column in a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_tablecolumn_set_formula.rs
    /// #
    /// # use rust_xlsxwriter::{Table, TableColumn, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the columns.
    ///     let mut table = Table::new();
    ///
    ///     // Add a structured reference formula to the last column and set the header
    ///     // caption.
    ///     let columns = vec![
    ///         TableColumn::new().set_header("Product"),
    ///         TableColumn::new().set_header("Quarter 1"),
    ///         TableColumn::new().set_header("Quarter 2"),
    ///         TableColumn::new().set_header("Quarter 3"),
    ///         TableColumn::new().set_header("Quarter 4"),
    ///         TableColumn::new()
    ///             .set_header("Totals")
    ///             .set_formula("SUM(Table1[@[Quarter 1]:[Quarter 4]])"),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 6, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/tablecolumn_set_formula.png">
    ///
    pub fn set_formula(mut self, formula: impl Into<Formula>) -> TableColumn {
        let mut formula = formula.into();
        formula = formula.clone().use_table_functions();
        self.formula = Some(formula);
        self
    }

    /// Set the format for a table column.
    ///
    /// It is sometimes required to format the data in the columns of a table.
    /// This can be done using the standard
    /// [`worksheet.write_with_format()`](Worksheet.write_with_format)) method
    /// but format can also be applied separately using
    /// `TableColumn.set_format()`.
    ///
    /// The most common format property to set for a table column is the [number
    /// format](Format::set_num_format), see the example below.
    ///
    /// # Parameters
    ///
    /// * `format` - The [`Format`] property for the data cells in the column.
    ///
    /// # Examples
    ///
    /// Example of adding a format to a column in a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_tablecolumn_set_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Table, TableColumn, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the columns.
    ///     let mut table = Table::new();
    ///
    ///     // Create a number format for number columns in the table.
    ///     let format = Format::new().set_num_format("$#,##0.00");
    ///
    ///     // Add a format to the number/currency columns.
    ///     let columns = vec![
    ///         TableColumn::new().set_header("Product"),
    ///         TableColumn::new().set_header("Q1").set_format(&format),
    ///         TableColumn::new().set_header("Q2").set_format(&format),
    ///         TableColumn::new().set_header("Q3").set_format(&format),
    ///         TableColumn::new().set_header("Q4").set_format(&format),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    ///
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/tablecolumn_set_format.png">
    ///
    pub fn set_format(mut self, format: &Format) -> TableColumn {
        self.format = Some(format.clone());
        self
    }

    /// Set the format for the header of the table column.
    ///
    /// The `set_header_format` method can be used to set the format for the
    /// column header in a worksheet table.
    ///
    /// # Parameters
    ///
    /// * `format` - The [`Format`] property for the column header.
    ///
    /// # Examples
    ///
    /// Example of adding a header format to a column in a worksheet table.
    ///
    /// ```
    /// # // This code is available in examples/doc_tablecolumn_set_header_format.rs
    /// #
    /// # use rust_xlsxwriter::{Format, Table, TableColumn, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data for the table.
    /// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
    /// #     let data = [
    /// #         [10000, 5000, 8000, 6000],
    /// #         [2000, 3000, 4000, 5000],
    /// #         [6000, 6000, 6500, 6000],
    /// #         [500, 300, 200, 700],
    /// #     ];
    /// #
    /// #     // Write the table data.
    /// #     worksheet.write_column(3, 1, items)?;
    /// #     worksheet.write_row_matrix(3, 2, data)?;
    /// #
    /// #     // Set the columns widths for clarity.
    /// #     for col_num in 1..=6u16 {
    /// #         worksheet.set_column_width(col_num, 12)?;
    /// #     }
    /// #
    ///     // Create a new table and configure the columns.
    ///     let mut table = Table::new();
    ///
    ///     // Create formats for the columns headers.
    ///     let format1 = Format::new().set_font_color("#FF0000");
    ///     let format2 = Format::new().set_font_color("#00FF00");
    ///     let format3 = Format::new().set_font_color("#0000FF");
    ///     let format4 = Format::new().set_font_color("#FFFF00");
    ///
    ///     // Add a format to the columns headers.
    ///     let columns = vec![
    ///         TableColumn::new().set_header("Product"),
    ///         TableColumn::new()
    ///             .set_header("Quarter 1")
    ///             .set_header_format(&format1),
    ///         TableColumn::new()
    ///             .set_header("Quarter 2")
    ///             .set_header_format(&format2),
    ///         TableColumn::new()
    ///             .set_header("Quarter 3")
    ///             .set_header_format(&format3),
    ///         TableColumn::new()
    ///             .set_header("Quarter 4")
    ///             .set_header_format(&format4),
    ///     ];
    ///     table.set_columns(&columns);
    ///
    ///     // Add the table to the worksheet.
    ///     worksheet.add_table(2, 1, 6, 5, &table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("tables.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/tablecolumn_set_header_format.png">
    ///
    pub fn set_header_format(mut self, format: &Format) -> TableColumn {
        self.header_format = Some(format.clone());
        self
    }

    // Convert the SUBTOTAL type to a worksheet formula.
    pub(crate) fn total_function(&self) -> Formula {
        let column_name = self
            .name
            .replace('\'', "''")
            .replace('#', "'#")
            .replace(']', "']")
            .replace('[', "'[");

        match self.total_function {
            TableFunction::None => Formula::new(""),
            TableFunction::Max => Formula::new(format!("SUBTOTAL(104,[{column_name}])")),
            TableFunction::Min => Formula::new(format!("SUBTOTAL(105,[{column_name}])")),
            TableFunction::Sum => Formula::new(format!("SUBTOTAL(109,[{column_name}])")),
            TableFunction::Var => Formula::new(format!("SUBTOTAL(110,[{column_name}])")),
            TableFunction::Count => Formula::new(format!("SUBTOTAL(103,[{column_name}])")),
            TableFunction::StdDev => Formula::new(format!("SUBTOTAL(107,[{column_name}])")),
            TableFunction::Average => Formula::new(format!("SUBTOTAL(101,[{column_name}])")),
            TableFunction::CountNumbers => Formula::new(format!("SUBTOTAL(102,[{column_name}])")),
        }
    }
}

impl Default for TableColumn {
    fn default() -> Self {
        Self::new()
    }
}

/// The `TableFunction` enum defines functions for worksheet table total rows.
///
/// The `TableFunction` enum contains definitions for the standard Excel
/// functions that are available via the dropdown in the total row of an Excel
/// table.
///
/// Excel also supports custom functions. These aren't currently supported by
/// `rust_xlsxwriter` but will be added in an upcoming release.
///
/// # Examples
///
/// Example of turning on the totals row at the bottom of a worksheet table with
/// subtotal functions.
///
/// ```
/// # // This code is available in examples/doc_table_set_total_row2.rs
/// #
/// # use rust_xlsxwriter::{Table, TableColumn, TableFunction, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Some sample data for the table.
/// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
/// #     let data = [
/// #         [10000, 5000, 8000, 6000],
/// #         [2000, 3000, 4000, 5000],
/// #         [6000, 6000, 6500, 6000],
/// #         [500, 300, 200, 700],
/// #     ];
/// #
/// #     // Write the table data.
/// #     worksheet.write_column(3, 1, items)?;
/// #     worksheet.write_row_matrix(3, 2, data)?;
/// #
/// #     // Set the columns widths for clarity.
/// #     for col_num in 1..=6u16 {
/// #         worksheet.set_column_width(col_num, 12)?;
/// #     }
/// #
///     // Create a new table and configure the total row.
///     let mut table = Table::new();
///     table.set_total_row(true);
///
///     // Set the caption and subtotal in the total row.
///     let columns = vec![
///         TableColumn::new().set_total_label("Totals"),
///         TableColumn::new().set_total_function(TableFunction::Sum),
///         TableColumn::new().set_total_function(TableFunction::Sum),
///         TableColumn::new().set_total_function(TableFunction::Sum),
///         TableColumn::new().set_total_function(TableFunction::Sum),
///     ];
///     table.set_columns(&columns);
///
///     // Add the table to the worksheet.
///     worksheet.add_table(2, 1, 7, 5, &table)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("tables.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/table_set_total_row2.png">
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TableFunction {
    /// The "total row" option is enable but there is no total function.
    None,

    /// Use the average function as the table total.
    Average,

    /// Use the count function as the table total.
    Count,

    /// Use the count numbers function as the table total.
    CountNumbers,

    /// Use the max function as the table total.
    Max,

    /// Use the min function as the table total.
    Min,

    /// Use the sum function as the table total.
    Sum,

    /// Use the standard deviation function as the table total.
    StdDev,

    /// Use the var function as the table total.
    Var,
}

impl fmt::Display for TableFunction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TableFunction::Max => write!(f, "max"),
            TableFunction::Min => write!(f, "min"),
            TableFunction::Sum => write!(f, "sum"),
            TableFunction::Var => write!(f, "var"),
            TableFunction::None => write!(f, "None"),
            TableFunction::Count => write!(f, "count"),
            TableFunction::StdDev => write!(f, "stdDev"),
            TableFunction::Average => write!(f, "average"),
            TableFunction::CountNumbers => write!(f, "countNums"),
        }
    }
}

/// The `TableStyle` enum defines the worksheet table styles.
///
/// Excel supports 61 different styles for tables divided into Light, Medium and
/// Dark categories. You can set one of these styles using a [`TableStyle`] enum
/// value.
///
/// <img src="https://rustxlsxwriter.github.io/images/table_styles.png">
///
/// The style is set via the [`table.set_style`](Table::set_style) method. The
/// default table style in Excel is equivalent to
/// [`TableStyle::Medium9`](TableStyle::Medium9).
///
/// # Examples
///
/// Example of setting the style of a worksheet table.
///
/// ```
/// # // This code is available in examples/doc_table_set_style.rs
/// #
/// # use rust_xlsxwriter::{Table, TableStyle, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Some sample data for the table.
/// #     let items = ["Apples", "Pears", "Bananas", "Oranges"];
/// #     let data = [
/// #         [10000, 5000, 8000, 6000],
/// #         [2000, 3000, 4000, 5000],
/// #         [6000, 6000, 6500, 6000],
/// #         [500, 300, 200, 700],
/// #     ];
/// #
/// #     // Write the table data.
/// #     worksheet.write_column(3, 1, items)?;
/// #     worksheet.write_row_matrix(3, 2, data)?;
/// #
/// #     // Set the columns widths for clarity.
/// #     for col_num in 1..=6u16 {
/// #         worksheet.set_column_width(col_num, 12)?;
/// #     }
/// #
///     // Create a new table and set the style.
///     let mut table = Table::new();
///     table.set_style(TableStyle::Medium10);
///
///     // Add the table to the worksheet.
///     worksheet.add_table(2, 1, 6, 5, &table)?;
///
/// #     // Save the file to disk.
/// #     workbook.save("tables.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/table_set_style.png">
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TableStyle {
    /// No table style.
    None,

    /// Table Style Light 1, White.
    Light1,

    /// Table Style Light 2, Light Blue.
    Light2,

    /// Table Style Light 3, Light Orange.
    Light3,

    /// Table Style Light 4, White.
    Light4,

    /// Table Style Light 5, Light Yellow.
    Light5,

    /// Table Style Light 6, Light Blue.
    Light6,

    /// Table Style Light 7, Light Green.
    Light7,

    /// Table Style Light 8, White.
    Light8,

    /// Table Style Light 9, Blue.
    Light9,

    /// Table Style Light 10, Orange.
    Light10,

    /// Table Style Light 11, White.
    Light11,

    /// Table Style Light 12, Gold.
    Light12,

    /// Table Style Light 13, Blue.
    Light13,

    /// Table Style Light 14, Green.
    Light14,

    /// Table Style Light 15, White.
    Light15,

    /// Table Style Light 16, Light Blue.
    Light16,

    /// Table Style Light 17, Light Orange.
    Light17,

    /// Table Style Light 18, White.
    Light18,

    /// Table Style Light 19, Light Yellow.
    Light19,

    /// Table Style Light 20, Light Blue.
    Light20,

    /// Table Style Light 21, Light Green.
    Light21,

    /// Table Style Medium 1, White.
    Medium1,

    /// Table Style Medium 2, Blue.
    Medium2,

    /// Table Style Medium 3, Orange.
    Medium3,

    /// Table Style Medium 4, White.
    Medium4,

    /// Table Style Medium 5, Gold.
    Medium5,

    /// Table Style Medium 6, Blue.
    Medium6,

    /// Table Style Medium 7, Green.
    Medium7,

    /// Table Style Medium 8, Light Grey.
    Medium8,

    /// Table Style Medium 9, Blue.
    Medium9,

    /// Table Style Medium 10, Orange.
    Medium10,

    /// Table Style Medium 11, Light Grey.
    Medium11,

    /// Table Style Medium 12, Gold.
    Medium12,

    /// Table Style Medium 13, Blue.
    Medium13,

    /// Table Style Medium 14, Green.
    Medium14,

    /// Table Style Medium 15, White.
    Medium15,

    /// Table Style Medium 16, Blue.
    Medium16,

    /// Table Style Medium 17, Orange.
    Medium17,

    /// Table Style Medium 18, White.
    Medium18,

    /// Table Style Medium 19, Gold.
    Medium19,

    /// Table Style Medium 20, Blue.
    Medium20,

    /// Table Style Medium 21, Green.
    Medium21,

    /// Table Style Medium 22, Light Grey.
    Medium22,

    /// Table Style Medium 23, Light Blue.
    Medium23,

    /// Table Style Medium 24, Light Orange.
    Medium24,

    /// Table Style Medium 25, Light Grey.
    Medium25,

    /// Table Style Medium 26, Light Yellow.
    Medium26,

    /// Table Style Medium 27, Light Blue.
    Medium27,

    /// Table Style Medium 28, Light Green.
    Medium28,

    /// Table Style Dark 1, Dark Grey.
    Dark1,

    /// Table Style Dark 2, Dark Blue.
    Dark2,

    /// Table Style Dark 3, Brown.
    Dark3,

    /// Table Style Dark 4, Grey.
    Dark4,

    /// Table Style Dark 5, Dark Yellow.
    Dark5,

    /// Table Style Dark 6, Blue.
    Dark6,

    /// Table Style Dark 7, Dark Green.
    Dark7,

    /// Table Style Dark 8, Light Grey.
    Dark8,

    /// Table Style Dark 9, Light Orange.
    Dark9,

    /// Table Style Dark 10, Gold.
    Dark10,

    /// Table Style Dark 11, Green.
    Dark11,
}

impl fmt::Display for TableStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TableStyle::None => write!(f, "TableStyleNone"),
            TableStyle::Light1 => write!(f, "TableStyleLight1"),
            TableStyle::Light2 => write!(f, "TableStyleLight2"),
            TableStyle::Light3 => write!(f, "TableStyleLight3"),
            TableStyle::Light4 => write!(f, "TableStyleLight4"),
            TableStyle::Light5 => write!(f, "TableStyleLight5"),
            TableStyle::Light6 => write!(f, "TableStyleLight6"),
            TableStyle::Light7 => write!(f, "TableStyleLight7"),
            TableStyle::Light8 => write!(f, "TableStyleLight8"),
            TableStyle::Light9 => write!(f, "TableStyleLight9"),
            TableStyle::Light10 => write!(f, "TableStyleLight10"),
            TableStyle::Light11 => write!(f, "TableStyleLight11"),
            TableStyle::Light12 => write!(f, "TableStyleLight12"),
            TableStyle::Light13 => write!(f, "TableStyleLight13"),
            TableStyle::Light14 => write!(f, "TableStyleLight14"),
            TableStyle::Light15 => write!(f, "TableStyleLight15"),
            TableStyle::Light16 => write!(f, "TableStyleLight16"),
            TableStyle::Light17 => write!(f, "TableStyleLight17"),
            TableStyle::Light18 => write!(f, "TableStyleLight18"),
            TableStyle::Light19 => write!(f, "TableStyleLight19"),
            TableStyle::Light20 => write!(f, "TableStyleLight20"),
            TableStyle::Light21 => write!(f, "TableStyleLight21"),
            TableStyle::Medium1 => write!(f, "TableStyleMedium1"),
            TableStyle::Medium2 => write!(f, "TableStyleMedium2"),
            TableStyle::Medium3 => write!(f, "TableStyleMedium3"),
            TableStyle::Medium4 => write!(f, "TableStyleMedium4"),
            TableStyle::Medium5 => write!(f, "TableStyleMedium5"),
            TableStyle::Medium6 => write!(f, "TableStyleMedium6"),
            TableStyle::Medium7 => write!(f, "TableStyleMedium7"),
            TableStyle::Medium8 => write!(f, "TableStyleMedium8"),
            TableStyle::Medium9 => write!(f, "TableStyleMedium9"),
            TableStyle::Medium10 => write!(f, "TableStyleMedium10"),
            TableStyle::Medium11 => write!(f, "TableStyleMedium11"),
            TableStyle::Medium12 => write!(f, "TableStyleMedium12"),
            TableStyle::Medium13 => write!(f, "TableStyleMedium13"),
            TableStyle::Medium14 => write!(f, "TableStyleMedium14"),
            TableStyle::Medium15 => write!(f, "TableStyleMedium15"),
            TableStyle::Medium16 => write!(f, "TableStyleMedium16"),
            TableStyle::Medium17 => write!(f, "TableStyleMedium17"),
            TableStyle::Medium18 => write!(f, "TableStyleMedium18"),
            TableStyle::Medium19 => write!(f, "TableStyleMedium19"),
            TableStyle::Medium20 => write!(f, "TableStyleMedium20"),
            TableStyle::Medium21 => write!(f, "TableStyleMedium21"),
            TableStyle::Medium22 => write!(f, "TableStyleMedium22"),
            TableStyle::Medium23 => write!(f, "TableStyleMedium23"),
            TableStyle::Medium24 => write!(f, "TableStyleMedium24"),
            TableStyle::Medium25 => write!(f, "TableStyleMedium25"),
            TableStyle::Medium26 => write!(f, "TableStyleMedium26"),
            TableStyle::Medium27 => write!(f, "TableStyleMedium27"),
            TableStyle::Medium28 => write!(f, "TableStyleMedium28"),
            TableStyle::Dark1 => write!(f, "TableStyleDark1"),
            TableStyle::Dark2 => write!(f, "TableStyleDark2"),
            TableStyle::Dark3 => write!(f, "TableStyleDark3"),
            TableStyle::Dark4 => write!(f, "TableStyleDark4"),
            TableStyle::Dark5 => write!(f, "TableStyleDark5"),
            TableStyle::Dark6 => write!(f, "TableStyleDark6"),
            TableStyle::Dark7 => write!(f, "TableStyleDark7"),
            TableStyle::Dark8 => write!(f, "TableStyleDark8"),
            TableStyle::Dark9 => write!(f, "TableStyleDark9"),
            TableStyle::Dark10 => write!(f, "TableStyleDark10"),
            TableStyle::Dark11 => write!(f, "TableStyleDark11"),
        }
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::table::Table;
    use crate::test_functions::xml_to_vec;
    use crate::{TableColumn, TableFunction, Worksheet, XlsxError};
    use pretty_assertions::assert_eq;

    #[test]
    fn test_row_methods() {
        let mut table = Table::new();
        table.cell_range.first_row = 0;
        table.cell_range.first_col = 0;
        table.cell_range.last_row = 9;
        table.cell_range.last_col = 4;

        assert_eq!(1, table.first_data_row());
        assert_eq!(9, table.last_data_row());

        table.set_total_row(true);
        assert_eq!(1, table.first_data_row());
        assert_eq!(8, table.last_data_row());

        table.set_header_row(false);
        assert_eq!(0, table.first_data_row());
        assert_eq!(8, table.last_data_row());
    }

    #[test]
    fn test_column_validation() {
        // Test the table column validation and checks.
        let mut table = Table::new();
        let default_headers = vec![
            String::from("Column1"),
            String::from("Column2"),
            String::from("Column3"),
            String::from("Column4"),
        ];

        table.cell_range.first_row = 0;
        table.cell_range.first_col = 0;
        table.cell_range.last_row = 8;
        table.cell_range.last_col = 4;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("foo"), // Lowercase duplicate.
        ];

        table.set_columns(&columns);
        let result = table.initialize_columns(&default_headers);

        assert!(matches!(result, Err(XlsxError::TableError(_))));

        let columns = vec![
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("column1"), // Lowercase duplicate.
        ];

        table.set_columns(&columns);
        let result = table.initialize_columns(&default_headers);

        assert!(matches!(result, Err(XlsxError::TableError(_))));
    }

    #[test]
    fn test_assemble1() {
        let mut table = Table::new();
        let default_headers = vec![
            String::from("Column1"),
            String::from("Column2"),
            String::from("Column3"),
            String::from("Column4"),
        ];

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble2() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 3;
        table.cell_range.first_col = 3;
        table.cell_range.last_row = 14;
        table.cell_range.last_col = 8;
        table.index = 2;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_style(crate::TableStyle::Light17);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Table2" displayName="Table2" ref="D4:I15" totalsRowShown="0">
                <autoFilter ref="D4:I15"/>
                <tableColumns count="6">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                    <tableColumn id="5" name="Column5"/>
                    <tableColumn id="6" name="Column6"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleLight17" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble3() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 4;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 15;
        table.cell_range.last_col = 3;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_first_column(true);
        table.set_last_column(true);
        table.set_banded_rows(false);
        table.set_banded_columns(true);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C5:D16" totalsRowShown="0">
                <autoFilter ref="C5:D16"/>
                <tableColumns count="2">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="1" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble4() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_autofilter(false);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble5() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 3;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_header_row(false);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C4:F13" headerRowCount="0" totalsRowShown="0">
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble6() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("Baz"),
        ];

        table.set_columns(&columns);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Foo"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Baz"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);

        // Try a variation with too many columns settings. It should be
        // truncated to the actual number of columns.
        let mut table = Table::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        let columns = vec![
            TableColumn::new().set_header("Foo"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_header("Baz"),
            TableColumn::new().set_header("Too many"),
        ];

        table.set_columns(&columns);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble7() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 13;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_total_row(true);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble8() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 2;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 13;
        table.cell_range.last_col = 5;
        table.index = 1;

        let columns = vec![
            TableColumn::new().set_total_label("Total"),
            TableColumn::default(),
            TableColumn::default(),
            TableColumn::new().set_total_function(TableFunction::Count),
        ];

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_columns(&columns);
        table.set_total_row(true);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
                <autoFilter ref="C3:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble9() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 1;
        table.cell_range.first_col = 1;
        table.cell_range.last_row = 7;
        table.cell_range.last_col = 10;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        let columns = vec![
            TableColumn::new()
                .set_total_label("Total")
                .set_total_function(TableFunction::Max), // Max should be ignore due to label.
            TableColumn::new().set_total_function(TableFunction::None), // Should be ignored.
            TableColumn::new().set_total_function(TableFunction::Average),
            TableColumn::new().set_total_function(TableFunction::Count),
            TableColumn::new().set_total_function(TableFunction::CountNumbers),
            TableColumn::new().set_total_function(TableFunction::Max),
            TableColumn::new().set_total_function(TableFunction::Min),
            TableColumn::new().set_total_function(TableFunction::Sum),
            TableColumn::new().set_total_function(TableFunction::StdDev),
            TableColumn::new().set_total_function(TableFunction::Var),
        ];

        table.set_columns(&columns);
        table.set_total_row(true);

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:K8" totalsRowCount="1">
                <autoFilter ref="B2:K7"/>
                <tableColumns count="10">
                    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3" totalsRowFunction="average"/>
                    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
                    <tableColumn id="5" name="Column5" totalsRowFunction="countNums"/>
                    <tableColumn id="6" name="Column6" totalsRowFunction="max"/>
                    <tableColumn id="7" name="Column7" totalsRowFunction="min"/>
                    <tableColumn id="8" name="Column8" totalsRowFunction="sum"/>
                    <tableColumn id="9" name="Column9" totalsRowFunction="stdDev"/>
                    <tableColumn id="10" name="Column10" totalsRowFunction="var"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_assemble10() {
        let mut table = Table::new();
        let worksheet = Worksheet::new();

        table.cell_range.first_row = 1;
        table.cell_range.first_col = 2;
        table.cell_range.last_row = 12;
        table.cell_range.last_col = 5;
        table.index = 1;

        let default_headers = worksheet.default_table_headers(
            table.cell_range.first_row,
            table.cell_range.first_col,
            table.cell_range.last_col,
            table.show_header_row,
        );

        table.set_name("MyTable");

        table.initialize_columns(&default_headers).unwrap();
        table.assemble_xml_file();

        let got = table.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="MyTable" displayName="MyTable" ref="C2:F13" totalsRowShown="0">
                <autoFilter ref="C2:F13"/>
                <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                </tableColumns>
                <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
            "#,
        );

        assert_eq!(expected, got);
    }
}
