// pivot_table - A module for creating the Excel pivotTable.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::fmt;
use std::io::Cursor;

use crate::utility;
use crate::worksheet::{COL_MAX, ROW_MAX};
use crate::xmlwriter::{
    escape_attribute_escapes, xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag,
    xml_start_tag,
};
use crate::{ColNum, Format, RowNum, XlsxError};

/// The `PivotTable` struct represents a worksheet pivot table.
///
/// Pivot tables in Excel are a way of summarizing, grouping and analyzing a
/// range of source data. The unique values of the source data fields can be
/// arranged as row and column labels, and the numeric fields are summarized,
/// for example as a sum or an average, for each label. Pivot tables can also
/// have filter fields, different layouts and formatting styles.
///
/// <img src="https://rustxlsxwriter.github.io/images/pivot_table_demo.png">
///
/// A pivot table is added to a worksheet via the
/// [`Worksheet::add_pivot_table()`](crate::Worksheet::add_pivot_table) method.
/// The source data range and fields of a pivot table are configured via a
/// `PivotTable` struct, but the source data is added using standard
/// [`Worksheet::write()`](crate::Worksheet::write) methods. Here is an example
/// that generates the pivot table shown above.
///
/// ```
/// use rust_xlsxwriter::{
///     PivotTable, PivotTableFunction, PivotTableLayout, PivotTableStyle, PivotTableValueField,
///     Workbook, XlsxError,
/// };
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet for the source data.
///     let worksheet = workbook.add_worksheet().set_name("Data")?;
///
///     // Some sample data to summarize.
///     let data = [
///         ("East", "Apple", 9000, "July"),
///         ("West", "Pear", 5000, "July"),
///         ("North", "Apple", 3000, "August"),
///         ("East", "Grape", 8000, "August"),
///         ("West", "Apple", 6000, "September"),
///         ("North", "Pear", 4000, "September"),
///         ("East", "Pear", 7000, "October"),
///         ("South", "Grape", 2000, "October"),
///     ];
///
///     // Write the header row and the data.
///     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
///     for (row, (region, item, volume, month)) in data.iter().enumerate() {
///         let row = row as u32 + 1;
///         worksheet.write(row, 0, *region)?;
///         worksheet.write(row, 1, *item)?;
///         worksheet.write(row, 2, *volume)?;
///         worksheet.write(row, 3, *month)?;
///     }
///
///     // Create a pivot table that summarizes the "Volume" by "Region" and "Item",
///     // with a filter on "Month".
///     let pivot_table = PivotTable::new()
///         .set_data_source("Data!A1:D9")
///         .add_row_field("Region")
///         .add_column_field("Item")
///         .add_filter_field("Month")
///         .add_value_field(
///             &PivotTableValueField::new("Volume")
///                 .set_function(PivotTableFunction::Sum)
///                 .set_custom_name("Total Volume")
///                 .set_num_format("#,##0"),
///         )
///         .set_style(PivotTableStyle::Medium4)
///         .set_layout(PivotTableLayout::Tabular);
///
///     // Add a worksheet for the pivot table and add the pivot table to it.
///     let worksheet = workbook.add_worksheet().set_name("Pivot")?;
///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
///
///     // Save the file to disk.
///     workbook.save("pivot_table.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// For more information on pivot tables see the Microsoft documentation on
/// [Overview of PivotTables and PivotCharts].
///
/// [Overview of PivotTables and PivotCharts]:
///     https://support.microsoft.com/en-us/office/overview-of-pivottables-and-pivotcharts-527c8fa3-02c0-445a-a2db-7794676bce96
///
/// # Limitations
///
/// Pivot tables in Excel use a cache of the source data. This data is filtered
/// and summarized to create the pivot table. The `rust_xlsxwriter` library
/// doesn't store the cached data, so the pivot table is written without its
/// data. When Excel opens the file, it will build the cache and the pivot table
/// from the source data. This leads to the following limitations:
///
/// - Applications that don't refresh pivot tables when the file is opened will
///   show an empty pivot table. This also applies to pivot tables on protected
///   worksheets.
/// - Excel will ask the end user to save the file even if no visible changes
///   were made.
/// - The size of the pivot table isn't known until Excel builds it, so only
///   the top left cells of the pivot table are checked for overlaps with
///   worksheet tables, merged ranges, other pivot tables and the source data.
///   Excel will show an error if the full pivot table overlaps them.
///
/// The overall limitation can't easily be fixed without replicating most of the
/// logic of the Excel Pivot Table engine.
///
#[derive(Clone)]
pub struct PivotTable {
    pub(crate) writer: Cursor<Vec<u8>>,

    pub(crate) name: String,
    pub(crate) cache_id: u32,
    pub(crate) source: PivotTableSource,
    source_error: Option<String>,
    pub(crate) field_names: Vec<String>,

    pub(crate) row: RowNum,
    pub(crate) col: ColNum,

    row_fields: Vec<String>,
    column_fields: Vec<String>,
    filter_fields: Vec<String>,
    pub(crate) value_fields: Vec<PivotTableValueField>,

    style: PivotTableStyle,
    layout: PivotTableLayout,
    show_row_grand_totals: bool,
    show_column_grand_totals: bool,
}

impl PivotTable {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new `PivotTable` struct instance.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> PivotTable {
        let writer = Cursor::new(Vec::with_capacity(2048));

        PivotTable {
            writer,
            name: String::new(),
            cache_id: 0,
            source: PivotTableSource::default(),
            source_error: None,
            field_names: vec![],
            row: 0,
            col: 0,
            row_fields: vec![],
            column_fields: vec![],
            filter_fields: vec![],
            value_fields: vec![],
            style: PivotTableStyle::Light16,
            layout: PivotTableLayout::Compact,
            show_row_grand_totals: true,
            show_column_grand_totals: true,
        }
    }

    /// Set the worksheet range of the source data for the pivot table.
    ///
    /// The first row of the range should contain the field names. The data
    /// source is required since you can't create a pivot table without it. All
    /// of the other pivot table settings are optional.
    ///
    /// The range is checked when the pivot table is added to a worksheet with
    /// [`Worksheet::add_pivot_table()`](crate::Worksheet::add_pivot_table). The
    /// sheet name and header row are checked when the file is saved.
    ///
    /// # Parameters
    ///
    /// - `source`: The source data range. This can be a string like
    ///   `"Sheet1!A1:D9"` or a tuple like `("Sheet1", 0, 0, 8, 3)`. See
    ///   [`IntoPivotTableSource`] for details.
    ///
    /// # Examples
    ///
    /// Example of creating a simple pivot table and setting the data source.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_data_source.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region". Set the data
    ///     // source to point to the data on the first worksheet.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_set_data_source.png">
    ///
    /// The data source can also be specified as a tuple. This will generate the
    /// same results as the previous example.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_data_source_tuple.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region". Set the data
    ///     // source to point to the data on the first worksheet, using a tuple.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source(("Sheet1", 0, 0, 8, 3))
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_data_source<T>(mut self, source: T) -> PivotTable
    where
        T: IntoPivotTableSource,
    {
        let (sheet_name, first_row, first_col, last_row, last_col) =
            match source.new_pivot_table_source() {
                Ok(range) => range,
                Err(range_string) => {
                    self.source_error = Some(range_string);
                    return self;
                }
            };

        self.source_error = None;
        self.source = PivotTableSource {
            sheet_name,
            first_row,
            first_col,
            last_row,
            last_col,
            field_names: std::mem::take(&mut self.source.field_names),
        };

        self
    }

    /// Add a row field to the pivot table.
    ///
    /// The unique values of a row field are shown as labels down the left side
    /// of the pivot table, and any value fields are summarized for each label.
    /// More than one row field can be added to create a hierarchy, with each
    /// field nested under the previous one.
    ///
    /// A field can only be used once in the row, column and filter fields.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the field. This must match a header in the first
    ///   row of the source data.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table with a single row field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_row_field1.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table with a single row field and no other fields.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region");
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_add_row_field1.png">
    ///
    /// Example of creating a pivot table with two row fields. The second row
    /// field is nested under the first.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_row_field2.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region" and by "Item"
    ///     // within each region.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_row_field("Item")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_add_row_field2.png">
    ///
    pub fn add_row_field(mut self, name: &str) -> PivotTable {
        self.row_fields.push(name.to_string());
        self
    }

    /// Add a value field to the pivot table.
    ///
    /// A value field is the source data that is summarized in the body of the
    /// pivot table, for each of the row and column labels. The default
    /// summary is a sum. The function, name and number format of the value
    /// field can be set with [`PivotTableValueField`].
    ///
    /// More than one value field can be added and each one is shown in a
    /// separate column. The same source field can be used more than once with
    /// different functions.
    ///
    /// # Parameters
    ///
    /// - `value_field`: A [`PivotTableValueField`] struct reference.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table with a single value field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_value_field1.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region".
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_add_value_field1.png">
    ///
    /// Example of creating a pivot table with two value fields. The values are
    /// shown in separate columns.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_value_field2.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableFunction, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that shows the sum and the average of the "Volume"
    ///     // by "Region".
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume"))
    ///         .add_value_field(
    ///             &PivotTableValueField::new("Volume").set_function(PivotTableFunction::Average),
    ///         );
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_add_value_field2.png">
    ///
    pub fn add_value_field(mut self, value_field: &PivotTableValueField) -> PivotTable {
        self.value_fields.push(value_field.clone());
        self
    }

    /// Add a column field to the pivot table.
    ///
    /// The unique values of a column field are shown as labels across the top
    /// of the pivot table, and any value fields are summarized for each label.
    /// More than one column field can be added to create a hierarchy, with
    /// each field nested under the previous one.
    ///
    /// A field can only be used once in the row, column and filter fields.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the field. This must match a header in the first
    ///   row of the source data.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table with a row field and a column field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_column_field.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region" in the rows and
    ///     // by "Item" in the columns.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_column_field("Item")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_add_column_field.png">
    ///
    pub fn add_column_field(mut self, name: &str) -> PivotTable {
        self.column_fields.push(name.to_string());
        self
    }

    /// Add a filter field to the pivot table.
    ///
    /// A filter field is shown as a dropdown list above the pivot table that
    /// can be used to filter the data by the values of the field.
    ///
    /// The filter fields are displayed above the pivot table, with a blank row
    /// between them and the pivot table body. The position set in
    /// [`Worksheet::add_pivot_table()`](crate::Worksheet::add_pivot_table) is
    /// the top left cell of the filter fields.
    ///
    /// A field can only be used once in the row, column and filter fields.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the field. This must match a header in the first
    ///   row of the source data.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table with a filter field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_add_filter_field.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region", with a filter
    ///     // on "Month".
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_filter_field("Month")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_add_filter_field.png">
    ///
    pub fn add_filter_field(mut self, name: &str) -> PivotTable {
        self.filter_fields.push(name.to_string());
        self
    }

    /// Set the style of the pivot table.
    ///
    /// The style can be any of the built-in styles in the Excel "Pivot Table
    /// Styles" gallery. The default style is [`PivotTableStyle::Light16`].
    /// Use [`PivotTableStyle::None`] to turn off the style.
    ///
    /// # Parameters
    ///
    /// - `style`: A [`PivotTableStyle`] enum value.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table and setting its style.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_style.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableStyle, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region" and "Item", and
    ///     // set the pivot table style.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_column_field("Item")
    ///         .add_value_field(&PivotTableValueField::new("Volume"))
    ///         .set_style(PivotTableStyle::Dark18);
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_set_style.png">
    ///
    pub fn set_style(mut self, style: PivotTableStyle) -> PivotTable {
        self.style = style;
        self
    }

    /// Set the report layout of the pivot table.
    ///
    /// The layout mainly affects how the row fields are arranged when there is
    /// more than one row field. The available options are:
    ///
    /// - [`PivotTableLayout::Compact`], which shows the pivot table in compact
    ///   form, with the row fields in one column. This is the default.
    /// - [`PivotTableLayout::Outline`], which shows the pivot table in outline
    ///   form, with each row field in its own column and the items of the next
    ///   field below the item in the current field.
    /// - [`PivotTableLayout::Tabular`], which shows the pivot table in tabular
    ///   form, with each row field in its own column and the items of the next
    ///   field starting on the same row as the item in the current field.
    ///
    /// # Parameters
    ///
    /// - `layout`: A [`PivotTableLayout`] enum value.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table with the Outline layout. Each row
    /// field is in its own column and the items of the next field are below the
    /// item in the current field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_layout1.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableLayout, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region" and by "Item"
    ///     // within each region, using the Outline layout.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_row_field("Item")
    ///         .add_value_field(&PivotTableValueField::new("Volume"))
    ///         .set_layout(PivotTableLayout::Outline);
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_set_layout1.png">
    ///
    /// Example of creating a pivot table with the Tabular layout. Each row
    /// field is in its own column.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_layout2.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableLayout, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region" and by "Item"
    ///     // within each region, using the Tabular layout.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_row_field("Item")
    ///         .add_value_field(&PivotTableValueField::new("Volume"))
    ///         .set_layout(PivotTableLayout::Tabular);
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_set_layout2.png">
    ///
    pub fn set_layout(mut self, layout: PivotTableLayout) -> PivotTable {
        self.layout = layout;
        self
    }

    /// Turn on/off the grand totals for the pivot table rows.
    ///
    /// The row grand totals are displayed in a "Grand Total" column to the
    /// right of the pivot table. They are on by default.
    ///
    /// In Excel this is the "Rows" part of the "Grand Totals" options, such as
    /// "On for Rows Only". The column grand totals are set with
    /// [`PivotTable::set_column_grand_totals()`].
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table and turning off the row and column grand
    /// totals.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_grand_totals.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region" and "Item", and
    ///     // turn off the row and column grand totals.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_column_field("Item")
    ///         .add_value_field(&PivotTableValueField::new("Volume"))
    ///         .set_row_grand_totals(false)
    ///         .set_column_grand_totals(false);
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_set_grand_totals.png">
    ///
    pub fn set_row_grand_totals(mut self, enable: bool) -> PivotTable {
        self.show_row_grand_totals = enable;
        self
    }

    /// Turn on/off the grand totals for the pivot table columns.
    ///
    /// The column grand totals are displayed in a "Grand Total" row at the
    /// bottom of the pivot table. They are on by default.
    ///
    /// In Excel this is the "Columns" part of the "Grand Totals" options, such
    /// as "On for Columns Only". The row grand totals are set with
    /// [`PivotTable::set_row_grand_totals()`], see the example there.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn set_column_grand_totals(mut self, enable: bool) -> PivotTable {
        self.show_column_grand_totals = enable;
        self
    }

    /// Set the field names of the pivot table source data.
    ///
    /// The field names of a pivot table are the header names of the columns in
    /// the source data. These are normally read automatically from the top row
    /// of the area defined by [`PivotTable::set_data_source()`].
    ///
    /// However, if the data source is on a worksheet in [constant or low memory
    /// mode](crate::performance#constant-memory-mode) the header row isn't
    /// retained in memory, and the field names must be set explicitly with
    /// this method.
    ///
    /// Note: the field names must match the header row of the source data.
    ///
    /// # Parameters
    ///
    /// - `field_names`: The field names of the source data columns, in column
    ///   order.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table and setting the field names explicitly.
    /// This is generally only required when the data source is on a worksheet in
    /// constant or low memory mode.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_set_field_names.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table and set the field names explicitly.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .set_field_names(&["Region", "Item", "Volume", "Month"])
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/pivot_table_set_field_names.png">
    ///
    pub fn set_field_names(mut self, field_names: &[impl AsRef<str>]) -> PivotTable {
        self.source.field_names = field_names
            .iter()
            .map(|name| name.as_ref().to_string())
            .collect();

        self
    }

    // Validate the pivot table data source and fields:
    //
    // - The source must be set and have at least one data row.
    // - Any field names must match the number of source columns.
    // - A field can only be used once in the row, column and filter fields,
    //   but it can also be a value field.
    // - Custom value field names must be unique, ignoring case.
    //
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        if let Some(range) = &self.source_error {
            return Err(XlsxError::PivotTableError(format!(
                "The pivot table data source '{range}' isn't a valid range like 'Sheet1!A1:D9'"
            )));
        }

        let source = &self.source;

        if source.sheet_name.is_empty() {
            return Err(XlsxError::PivotTableError(
                "The pivot table data source must be set with set_data_source()".to_string(),
            ));
        }

        if source.last_row >= ROW_MAX || source.last_col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        if source.first_row > source.last_row || source.first_col > source.last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // The source data needs a header row and at least one row of data.
        if source.first_row == source.last_row {
            return Err(XlsxError::PivotTableError(
                "The pivot table data source must have a header row and at least one data row"
                    .to_string(),
            ));
        }

        // Any user defined field names must match the number of columns.
        if !source.field_names.is_empty()
            && source.field_names.len() != source.num_fields() as usize
        {
            return Err(XlsxError::PivotTableError(format!(
                "The number of field names ({}) doesn't match the number of data source columns ({})",
                source.field_names.len(),
                source.num_fields()
            )));
        }

        let mut seen_names = vec![];

        for name in self
            .row_fields
            .iter()
            .chain(&self.column_fields)
            .chain(&self.filter_fields)
        {
            if seen_names.contains(&name) {
                return Err(XlsxError::PivotTableError(format!(
                    "Field '{name}' is used more than once in the row, column or filter fields"
                )));
            }

            seen_names.push(name);
        }

        let value_names = self.value_field_names();
        for (value_field, name) in self.value_fields.iter().zip(&value_names) {
            if !value_field.custom_name.is_empty()
                && value_names
                    .iter()
                    .filter(|other| other.to_lowercase() == name.to_lowercase())
                    .count()
                    > 1
            {
                return Err(XlsxError::PivotTableError(format!(
                    "Value field name '{name}' is used more than once"
                )));
            }
        }

        Ok(())
    }

    // Validate that the pivot table field names match the field names in the
    // header row of the source data. This can only be done when the workbook
    // is saved since the header data may not have been written before then.
    // A custom value field name also can't be the same as a header name,
    // ignoring case.
    pub(crate) fn validate_field_names(&self) -> Result<(), XlsxError> {
        for name in self
            .row_fields
            .iter()
            .chain(&self.column_fields)
            .chain(&self.filter_fields)
            .chain(
                self.value_fields
                    .iter()
                    .map(|value_field| &value_field.field),
            )
        {
            if self.field_index(name).is_none() {
                return Err(XlsxError::PivotTableError(format!(
                    "Field '{name}' doesn't match a header in the source data"
                )));
            }
        }

        for value_field in &self.value_fields {
            let name = &value_field.custom_name;
            if !name.is_empty()
                && self
                    .field_names
                    .iter()
                    .any(|field_name| field_name.to_lowercase() == name.to_lowercase())
            {
                return Err(XlsxError::PivotTableError(format!(
                    "Value field name '{name}' is the same as a header in the source data"
                )));
            }
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the pivotTableDefinition element.
        self.write_pivot_table_definition();

        // Write the location element.
        self.write_location();

        // Write the pivotFields element.
        self.write_pivot_fields();

        let has_fields = !self.row_fields.is_empty()
            || !self.column_fields.is_empty()
            || !self.value_fields.is_empty();

        if !self.row_fields.is_empty() {
            // Write the rowFields element.
            self.write_axis_fields("rowFields", &self.row_fields.clone());
        }

        if has_fields && (self.row_fields.is_empty() || self.show_column_grand_totals) {
            // Write the rowItems element.
            self.write_axis_items("rowItems", !self.row_fields.is_empty());
        }

        // Multiple value fields are displayed in columns via a virtual "Values"
        // field that is added to the column fields.
        let has_values_field = self.value_fields.len() > 1;

        if !self.column_fields.is_empty() || has_values_field {
            // Write the colFields element.
            self.write_col_fields(has_values_field);
        }

        if has_fields {
            if has_values_field {
                // With column fields the value items are all grand totals.
                if self.column_fields.is_empty() || self.show_row_grand_totals {
                    // Write the colItems element.
                    self.write_value_col_items();
                }
            } else if self.column_fields.is_empty() || self.show_row_grand_totals {
                // Write the colItems element.
                self.write_axis_items("colItems", !self.column_fields.is_empty());
            }
        }

        if !self.filter_fields.is_empty() {
            // Write the pageFields element.
            self.write_page_fields();
        }

        if !self.value_fields.is_empty() {
            // Write the dataFields element.
            self.write_data_fields();
        }

        // Write the pivotTableStyleInfo element.
        self.write_pivot_table_style_info();

        // Close the pivotTableDefinition tag.
        xml_end_tag(&mut self.writer, "pivotTableDefinition");
    }

    // Write the <pivotTableDefinition> element.
    fn write_pivot_table_definition(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        let mut attributes = vec![
            ("xmlns", xmlns.to_string()),
            ("name", escape_attribute_escapes(&self.name).into_owned()),
            ("cacheId", self.cache_id.to_string()),
            ("applyNumberFormats", "0".to_string()),
            ("applyBorderFormats", "0".to_string()),
            ("applyFontFormats", "0".to_string()),
            ("applyPatternFormats", "0".to_string()),
            ("applyAlignmentFormats", "0".to_string()),
            ("applyWidthHeightFormats", "1".to_string()),
            ("dataCaption", "Values".to_string()),
            ("updatedVersion", "3".to_string()),
            ("minRefreshableVersion", "3".to_string()),
            ("showCalcMbrs", "0".to_string()),
            ("useAutoFormatting", "1".to_string()),
        ];

        // The Excel UI and XML names are reversed. The UI row grand totals are
        // the "Grand Total" column, which is the XML colGrandTotals, and vice
        // versa.
        if !self.show_column_grand_totals {
            attributes.push(("rowGrandTotals", "0".to_string()));
        }

        if !self.show_row_grand_totals {
            attributes.push(("colGrandTotals", "0".to_string()));
        }

        attributes.extend([
            ("itemPrintTitles", "1".to_string()),
            ("createdVersion", "3".to_string()),
            ("indent", "0".to_string()),
        ]);

        match self.layout {
            PivotTableLayout::Compact => {
                attributes.push(("outline", "1".to_string()));
                attributes.push(("outlineData", "1".to_string()));
            }
            PivotTableLayout::Outline => {
                attributes.push(("compact", "0".to_string()));
                attributes.push(("outline", "1".to_string()));
                attributes.push(("outlineData", "1".to_string()));
                attributes.push(("compactData", "0".to_string()));
            }
            PivotTableLayout::Tabular => {
                attributes.push(("compact", "0".to_string()));
                attributes.push(("compactData", "0".to_string()));
            }
        }

        attributes.push(("multipleFieldFilters", "0".to_string()));

        xml_start_tag(&mut self.writer, "pivotTableDefinition", &attributes);
    }

    // Write the <location> element.
    fn write_location(&mut self) {
        let mut first_header_row = 1;
        // There is a header row for each column field, including the virtual
        // "Values" field used for multiple value fields.
        let mut num_col_fields = self.column_fields.len();
        if self.value_fields.len() > 1 {
            num_col_fields += 1;
        }
        let mut first_data_row = 1 + num_col_fields;
        // In the compact layout the row fields share a column. In the outline
        // and tabular layouts each row field has its own column.
        let first_data_col = if self.row_fields.is_empty() && self.column_fields.is_empty() {
            0
        } else if matches!(self.layout, PivotTableLayout::Compact) {
            1
        } else {
            self.row_fields.len().max(1)
        };

        // A pivot table with only filter fields has no header or data rows.
        if self.row_fields.is_empty()
            && self.column_fields.is_empty()
            && self.value_fields.is_empty()
            && !self.filter_fields.is_empty()
        {
            first_header_row = 0;
            first_data_row = 0;
        }

        let num_filters = self.filter_fields.len() as RowNum;

        let mut attributes = vec![
            ("ref", utility::row_col_to_cell(self.body_row(), self.col)),
            ("firstHeaderRow", first_header_row.to_string()),
            ("firstDataRow", first_data_row.to_string()),
            ("firstDataCol", first_data_col.to_string()),
        ];

        if num_filters > 0 {
            attributes.push(("rowPageCount", num_filters.to_string()));
            attributes.push(("colPageCount", "1".to_string()));
        }

        xml_empty_tag(&mut self.writer, "location", &attributes);
    }

    // Write the <pivotFields> element.
    fn write_pivot_fields(&mut self) {
        let count = self.source.num_fields();
        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "pivotFields", &attributes);

        for index in 0..count as usize {
            // Write the pivotField element.
            self.write_pivot_field(index);
        }

        xml_end_tag(&mut self.writer, "pivotFields");
    }

    // Write the <pivotField> element.
    fn write_pivot_field(&mut self, index: usize) {
        let is_row_field = self
            .row_fields
            .iter()
            .any(|name| self.field_index(name) == Some(index));

        let is_column_field = self
            .column_fields
            .iter()
            .any(|name| self.field_index(name) == Some(index));

        let is_filter_field = self
            .filter_fields
            .iter()
            .any(|name| self.field_index(name) == Some(index));

        let is_data_field = self
            .value_fields
            .iter()
            .any(|value_field| self.field_index(&value_field.field) == Some(index));

        let mut attributes = vec![];

        if is_row_field {
            attributes.push(("axis", "axisRow"));
        } else if is_column_field {
            attributes.push(("axis", "axisCol"));
        } else if is_filter_field {
            attributes.push(("axis", "axisPage"));
        }

        if is_data_field {
            attributes.push(("dataField", "1"));
        }

        match self.layout {
            PivotTableLayout::Compact => {}
            PivotTableLayout::Outline => {
                attributes.push(("compact", "0"));
            }
            PivotTableLayout::Tabular => {
                attributes.push(("compact", "0"));
                attributes.push(("outline", "0"));
            }
        }

        attributes.push(("showAll", "0"));

        if is_row_field || is_column_field || is_filter_field {
            xml_start_tag(&mut self.writer, "pivotField", &attributes);

            // Write the items element.
            self.write_items();

            xml_end_tag(&mut self.writer, "pivotField");
        } else {
            xml_empty_tag(&mut self.writer, "pivotField", &attributes);
        }
    }

    // Write the <items> element. The items are generated by Excel when the
    // pivot table is refreshed so we only write the default item.
    fn write_items(&mut self) {
        let attributes = [("count", "1")];

        xml_start_tag(&mut self.writer, "items", &attributes);

        // Write the item element.
        let attributes = [("t", "default")];
        xml_empty_tag(&mut self.writer, "item", &attributes);

        xml_end_tag(&mut self.writer, "items");
    }

    // Write the <rowFields> or <colFields> element.
    fn write_axis_fields(&mut self, tag: &str, names: &[String]) {
        let attributes = [("count", names.len().to_string())];

        xml_start_tag(&mut self.writer, tag, &attributes);

        for name in names {
            if let Some(index) = self.field_index(name) {
                // Write the field element.
                let attributes = [("x", index.to_string())];
                xml_empty_tag(&mut self.writer, "field", &attributes);
            }
        }

        xml_end_tag(&mut self.writer, tag);
    }

    // Write the <colFields> element, with the virtual "Values" field (index
    // -2) after the column fields if there are multiple value fields.
    fn write_col_fields(&mut self, has_values_field: bool) {
        let mut indices: Vec<String> = self
            .column_fields
            .iter()
            .filter_map(|name| self.field_index(name))
            .map(|index| index.to_string())
            .collect();

        if has_values_field {
            indices.push("-2".to_string());
        }

        let attributes = [("count", indices.len().to_string())];

        xml_start_tag(&mut self.writer, "colFields", &attributes);

        for index in indices {
            // Write the field element.
            let attributes = [("x", index)];
            xml_empty_tag(&mut self.writer, "field", &attributes);
        }

        xml_end_tag(&mut self.writer, "colFields");
    }

    // Write the <colItems> element for multiple value fields. If there are no
    // column fields Excel writes an item for each value field. Otherwise we
    // write a grand total for each value field.
    fn write_value_col_items(&mut self) {
        let count = self.value_fields.len();
        let has_column_fields = !self.column_fields.is_empty();
        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "colItems", &attributes);

        for index in 0..count {
            let mut attributes = vec![];

            if has_column_fields {
                attributes.push(("t", "grand".to_string()));
            }

            if index > 0 {
                attributes.push(("i", index.to_string()));
            }

            // Write the i element.
            xml_start_tag(&mut self.writer, "i", &attributes);

            // Write the x element.
            if has_column_fields || index == 0 {
                xml_empty_tag_only(&mut self.writer, "x");
            } else {
                let attributes = [("v", index.to_string())];
                xml_empty_tag(&mut self.writer, "x", &attributes);
            }

            xml_end_tag(&mut self.writer, "i");
        }

        xml_end_tag(&mut self.writer, "colItems");
    }

    // Write the <rowItems> or <colItems> element. The items are generated by
    // Excel when the pivot table is refreshed so we only write the grand
    // total, if the axis has fields, or an empty item.
    fn write_axis_items(&mut self, tag: &str, has_fields: bool) {
        let attributes = [("count", "1")];

        xml_start_tag(&mut self.writer, tag, &attributes);

        // Write the i element.
        if has_fields {
            let attributes = [("t", "grand")];
            xml_start_tag(&mut self.writer, "i", &attributes);
            xml_empty_tag_only(&mut self.writer, "x");
            xml_end_tag(&mut self.writer, "i");
        } else {
            xml_empty_tag_only(&mut self.writer, "i");
        }

        xml_end_tag(&mut self.writer, tag);
    }

    // Write the <pageFields> element.
    fn write_page_fields(&mut self) {
        let attributes = [("count", self.filter_fields.len().to_string())];

        xml_start_tag(&mut self.writer, "pageFields", &attributes);

        for name in self.filter_fields.clone() {
            if let Some(index) = self.field_index(&name) {
                // Write the pageField element.
                let attributes = [("fld", index.to_string()), ("hier", "-1".to_string())];
                xml_empty_tag(&mut self.writer, "pageField", &attributes);
            }
        }

        xml_end_tag(&mut self.writer, "pageFields");
    }

    // Write the <dataFields> element.
    fn write_data_fields(&mut self) {
        let attributes = [("count", self.value_fields.len().to_string())];

        xml_start_tag(&mut self.writer, "dataFields", &attributes);

        let names = self.value_field_names();
        for (value_field, name) in self.value_fields.clone().iter().zip(names) {
            // Write the dataField element.
            self.write_data_field(value_field, &name);
        }

        xml_end_tag(&mut self.writer, "dataFields");
    }

    // Write the <dataField> element.
    fn write_data_field(&mut self, value_field: &PivotTableValueField, name: &str) {
        let Some(index) = self.field_index(&value_field.field) else {
            return;
        };

        let mut attributes = vec![
            ("name", escape_attribute_escapes(name).into_owned()),
            ("fld", index.to_string()),
        ];

        if !matches!(value_field.function, PivotTableFunction::Sum) {
            attributes.push(("subtotal", value_field.function.to_string()));
        }

        attributes.push(("baseField", "0".to_string()));
        attributes.push(("baseItem", "0".to_string()));

        if value_field.num_format_index > 0 {
            attributes.push(("numFmtId", value_field.num_format_index.to_string()));
        }

        xml_empty_tag(&mut self.writer, "dataField", &attributes);
    }

    // Get the first row of the pivot table body. The (row, col) position is the
    // top left of the pivot table including any filter fields. The body is
    // below the filter fields and a blank row.
    pub(crate) fn body_row(&self) -> RowNum {
        let num_filters = self.filter_fields.len() as RowNum;

        if num_filters == 0 {
            self.row
        } else {
            self.row + num_filters + 1
        }
    }

    // Check if a cell is in the known part of the pivot table: the anchor
    // column from the anchor down to the first body row.
    pub(crate) fn overlaps_anchor(&self, row: RowNum, col: ColNum) -> bool {
        col == self.col && (self.row..=self.body_row()).contains(&row)
    }

    // Get the display names of the value fields. These are either the custom
    // name or a default name like "Sum of Volume". Excel makes duplicate
    // default names unique by appending a number, such as "Sum of Volume2".
    // The names are compared ignoring case.
    fn value_field_names(&self) -> Vec<String> {
        let mut names: Vec<String> = vec![];
        let contains = |names: &[String], name: &str| {
            names
                .iter()
                .any(|other| other.to_lowercase() == name.to_lowercase())
        };

        for value_field in &self.value_fields {
            let mut name = if value_field.custom_name.is_empty() {
                format!(
                    "{} of {}",
                    value_field.function.caption(),
                    value_field.field
                )
            } else {
                value_field.custom_name.clone()
            };

            if value_field.custom_name.is_empty() && contains(&names, &name) {
                let mut suffix = 2;
                while contains(&names, &format!("{name}{suffix}")) {
                    suffix += 1;
                }
                name = format!("{name}{suffix}");
            }

            names.push(name);
        }

        names
    }

    // Get the index of a field in the source data from its header name.
    fn field_index(&self, name: &str) -> Option<usize> {
        self.field_names
            .iter()
            .position(|field_name| field_name == name)
    }

    // Write the <pivotTableStyleInfo> element.
    fn write_pivot_table_style_info(&mut self) {
        let mut attributes = vec![];

        if !matches!(self.style, PivotTableStyle::None) {
            attributes.push(("name", self.style.to_string()));
        }

        attributes.push(("showRowHeaders", "1".to_string()));
        attributes.push(("showColHeaders", "1".to_string()));
        attributes.push(("showRowStripes", "0".to_string()));
        attributes.push(("showColStripes", "0".to_string()));
        attributes.push(("showLastColumn", "1".to_string()));

        xml_empty_tag(&mut self.writer, "pivotTableStyleInfo", &attributes);
    }
}

// -----------------------------------------------------------------------
// PivotTableValueField
// -----------------------------------------------------------------------

/// The `PivotTableValueField` struct represents a value field in a pivot table.
///
/// It is passed as a parameter to the [`PivotTable::add_value_field()`] method.
///
/// # Examples
///
/// Example of creating a pivot table with three value fields. The values are
/// shown in separate columns.
///
/// ```
/// # // This code is available in examples/doc_pivot_table_add_value_field3.rs
/// #
/// # use rust_xlsxwriter::{PivotTable, PivotTableFunction, PivotTableValueField, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet for the source data.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Some sample data to summarize.
/// #     let data = [
/// #         ("East", "Apple", 9000, "July"),
/// #         ("West", "Pear", 5000, "July"),
/// #         ("North", "Apple", 3000, "August"),
/// #         ("East", "Grape", 8000, "August"),
/// #         ("West", "Apple", 6000, "September"),
/// #         ("North", "Pear", 4000, "September"),
/// #         ("East", "Pear", 7000, "October"),
/// #         ("South", "Grape", 2000, "October"),
/// #     ];
/// #
/// #     // Write the header row and the data.
/// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
/// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
/// #         let row = row as u32 + 1;
/// #         worksheet.write(row, 0, *region)?;
/// #         worksheet.write(row, 1, *item)?;
/// #         worksheet.write(row, 2, *volume)?;
/// #         worksheet.write(row, 3, *month)?;
/// #     }
/// #
///     // Create a pivot table that shows the sum, the average and the maximum of
///     // the "Volume" by "Region".
///     let pivot_table = PivotTable::new()
///         .set_data_source("Sheet1!A1:D9")
///         .add_row_field("Region")
///         .add_value_field(&PivotTableValueField::new("Volume"))
///         .add_value_field(
///             &PivotTableValueField::new("Volume").set_function(PivotTableFunction::Average),
///         )
///         .add_value_field(
///             &PivotTableValueField::new("Volume").set_function(PivotTableFunction::Max),
///         );
/// #
/// #     // Add a worksheet for the pivot table and add the pivot table to it.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.add_pivot_table(2, 0, &pivot_table)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("pivot_table.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/pivot_table_add_value_field3.png">
///
///
#[derive(Clone)]
pub struct PivotTableValueField {
    field: String,
    function: PivotTableFunction,
    custom_name: String,
    pub(crate) format: Option<Format>,
    pub(crate) num_format_index: u16,
}

impl PivotTableValueField {
    /// Create a new `PivotTableValueField` struct instance.
    ///
    /// # Parameters
    ///
    /// - `field`: The name of the source data field to summarize. This must
    ///   match a header in the first row of the source data.
    ///
    pub fn new(field: &str) -> PivotTableValueField {
        PivotTableValueField {
            field: field.to_string(),
            function: PivotTableFunction::Sum,
            custom_name: String::new(),
            format: None,
            num_format_index: 0,
        }
    }

    /// Set the function used to summarize the value field.
    ///
    /// The available functions are the options in the Excel "Summarize value
    /// field by" list, such as Sum, Count, Average, Max and Min. See
    /// [`PivotTableFunction`] for the full list. The default function is
    /// [`PivotTableFunction::Sum`].
    ///
    /// # Parameters
    ///
    /// - `function`: A [`PivotTableFunction`] enum value.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table and setting the function used to
    /// summarize a value field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_value_field_set_function.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableFunction, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that shows the average of the "Volume" by "Region".
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_value_field(
    ///             &PivotTableValueField::new("Volume").set_function(PivotTableFunction::Average),
    ///         );
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_value_field_set_function.png">
    ///
    pub fn set_function(mut self, function: PivotTableFunction) -> PivotTableValueField {
        self.function = function;
        self
    }

    /// Set a custom name for the value field.
    ///
    /// The default name is the function and field name, such as "Sum of
    /// Volume".
    ///
    /// Note: Excel doesn't allow the custom name to be the same as a header in
    /// the source data, or as another value field name, ignoring case.
    ///
    /// # Parameters
    ///
    /// - `name`: The custom name of the value field.
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table and setting a custom name for a value
    /// field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_value_field_set_custom_name.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region", with a custom
    ///     // name for the value field.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume").set_custom_name("Total Volume"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_value_field_set_custom_name.png">
    ///
    pub fn set_custom_name(mut self, name: impl Into<String>) -> PivotTableValueField {
        self.custom_name = name.into();
        self
    }

    /// Set the number format for the value field.
    ///
    /// # Parameters
    ///
    /// - `num_format`: The number format string, such as `"$#,##0"`. See
    ///   [`Format::set_num_format()`](crate::Format::set_num_format).
    ///
    /// # Examples
    ///
    /// Example of creating a pivot table and setting the number format of a
    /// value field.
    ///
    /// ```
    /// # // This code is available in examples/doc_pivot_table_value_field_set_num_format.rs
    /// #
    /// # use rust_xlsxwriter::{PivotTable, PivotTableValueField, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet for the source data.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Some sample data to summarize.
    /// #     let data = [
    /// #         ("East", "Apple", 9000, "July"),
    /// #         ("West", "Pear", 5000, "July"),
    /// #         ("North", "Apple", 3000, "August"),
    /// #         ("East", "Grape", 8000, "August"),
    /// #         ("West", "Apple", 6000, "September"),
    /// #         ("North", "Pear", 4000, "September"),
    /// #         ("East", "Pear", 7000, "October"),
    /// #         ("South", "Grape", 2000, "October"),
    /// #     ];
    /// #
    /// #     // Write the header row and the data.
    /// #     worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"])?;
    /// #     for (row, (region, item, volume, month)) in data.iter().enumerate() {
    /// #         let row = row as u32 + 1;
    /// #         worksheet.write(row, 0, *region)?;
    /// #         worksheet.write(row, 1, *item)?;
    /// #         worksheet.write(row, 2, *volume)?;
    /// #         worksheet.write(row, 3, *month)?;
    /// #     }
    /// #
    ///     // Create a pivot table that sums the "Volume" by "Region", with a currency
    ///     // number format for the value field.
    ///     let pivot_table = PivotTable::new()
    ///         .set_data_source("Sheet1!A1:D9")
    ///         .add_row_field("Region")
    ///         .add_value_field(&PivotTableValueField::new("Volume").set_num_format("$#,##0"));
    ///
    ///     // Add a worksheet for the pivot table and add the pivot table to it.
    ///     let worksheet = workbook.add_worksheet();
    ///     worksheet.add_pivot_table(2, 0, &pivot_table)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("pivot_table.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/pivot_table_value_field_set_num_format.png">
    ///
    pub fn set_num_format(mut self, num_format: impl Into<String>) -> PivotTableValueField {
        // The number format is stored internally as a cell Format so that it
        // shares the cell number format indices.
        self.format = Some(Format::new().set_num_format(num_format));
        self
    }
}

// -----------------------------------------------------------------------
// PivotTableLayout
// -----------------------------------------------------------------------

/// The `PivotTableLayout` enum defines the report layout of a pivot table.
///
/// These are the layouts in the Excel "Report Layout" menu. They are used by the
/// [`PivotTable::set_layout()`] method to define how the pivot table is
/// displayed.
///
#[derive(Clone)]
pub enum PivotTableLayout {
    /// Show the pivot table in compact form, with the row fields in one
    /// column. This is the default.
    Compact,

    /// Show the pivot table in outline form, with each row field in its own
    /// column and the items of the next field below the item in the current
    /// field.
    Outline,

    /// Show the pivot table in tabular form, with each row field in its own
    /// column and the items of the next field starting on the same row as the
    /// item in the current field.
    Tabular,
}

// -----------------------------------------------------------------------
// PivotTableStyle
// -----------------------------------------------------------------------

/// The `PivotTableStyle` enum defines the built-in Excel pivot table styles.
///
/// These are the styles in the Excel "Pivot Table Styles" gallery. They are
/// used by the [`PivotTable::set_style()`] method. The default style is
/// [`PivotTableStyle::Light16`]. Use [`PivotTableStyle::None`] to turn off the
/// style.
///
/// <img src="https://rustxlsxwriter.github.io/images/pivot_table_set_style.png">
///
#[derive(Clone)]
pub enum PivotTableStyle {
    /// No pivot table style.
    None,

    /// Pivot Style Light 1.
    Light1,

    /// Pivot Style Light 2.
    Light2,

    /// Pivot Style Light 3.
    Light3,

    /// Pivot Style Light 4.
    Light4,

    /// Pivot Style Light 5.
    Light5,

    /// Pivot Style Light 6.
    Light6,

    /// Pivot Style Light 7.
    Light7,

    /// Pivot Style Light 8.
    Light8,

    /// Pivot Style Light 9.
    Light9,

    /// Pivot Style Light 10.
    Light10,

    /// Pivot Style Light 11.
    Light11,

    /// Pivot Style Light 12.
    Light12,

    /// Pivot Style Light 13.
    Light13,

    /// Pivot Style Light 14.
    Light14,

    /// Pivot Style Light 15.
    Light15,

    /// Pivot Style Light 16.
    Light16,

    /// Pivot Style Light 17.
    Light17,

    /// Pivot Style Light 18.
    Light18,

    /// Pivot Style Light 19.
    Light19,

    /// Pivot Style Light 20.
    Light20,

    /// Pivot Style Light 21.
    Light21,

    /// Pivot Style Light 22.
    Light22,

    /// Pivot Style Light 23.
    Light23,

    /// Pivot Style Light 24.
    Light24,

    /// Pivot Style Light 25.
    Light25,

    /// Pivot Style Light 26.
    Light26,

    /// Pivot Style Light 27.
    Light27,

    /// Pivot Style Light 28.
    Light28,

    /// Pivot Style Medium 1.
    Medium1,

    /// Pivot Style Medium 2.
    Medium2,

    /// Pivot Style Medium 3.
    Medium3,

    /// Pivot Style Medium 4.
    Medium4,

    /// Pivot Style Medium 5.
    Medium5,

    /// Pivot Style Medium 6.
    Medium6,

    /// Pivot Style Medium 7.
    Medium7,

    /// Pivot Style Medium 8.
    Medium8,

    /// Pivot Style Medium 9.
    Medium9,

    /// Pivot Style Medium 10.
    Medium10,

    /// Pivot Style Medium 11.
    Medium11,

    /// Pivot Style Medium 12.
    Medium12,

    /// Pivot Style Medium 13.
    Medium13,

    /// Pivot Style Medium 14.
    Medium14,

    /// Pivot Style Medium 15.
    Medium15,

    /// Pivot Style Medium 16.
    Medium16,

    /// Pivot Style Medium 17.
    Medium17,

    /// Pivot Style Medium 18.
    Medium18,

    /// Pivot Style Medium 19.
    Medium19,

    /// Pivot Style Medium 20.
    Medium20,

    /// Pivot Style Medium 21.
    Medium21,

    /// Pivot Style Medium 22.
    Medium22,

    /// Pivot Style Medium 23.
    Medium23,

    /// Pivot Style Medium 24.
    Medium24,

    /// Pivot Style Medium 25.
    Medium25,

    /// Pivot Style Medium 26.
    Medium26,

    /// Pivot Style Medium 27.
    Medium27,

    /// Pivot Style Medium 28.
    Medium28,

    /// Pivot Style Dark 1.
    Dark1,

    /// Pivot Style Dark 2.
    Dark2,

    /// Pivot Style Dark 3.
    Dark3,

    /// Pivot Style Dark 4.
    Dark4,

    /// Pivot Style Dark 5.
    Dark5,

    /// Pivot Style Dark 6.
    Dark6,

    /// Pivot Style Dark 7.
    Dark7,

    /// Pivot Style Dark 8.
    Dark8,

    /// Pivot Style Dark 9.
    Dark9,

    /// Pivot Style Dark 10.
    Dark10,

    /// Pivot Style Dark 11.
    Dark11,

    /// Pivot Style Dark 12.
    Dark12,

    /// Pivot Style Dark 13.
    Dark13,

    /// Pivot Style Dark 14.
    Dark14,

    /// Pivot Style Dark 15.
    Dark15,

    /// Pivot Style Dark 16.
    Dark16,

    /// Pivot Style Dark 17.
    Dark17,

    /// Pivot Style Dark 18.
    Dark18,

    /// Pivot Style Dark 19.
    Dark19,

    /// Pivot Style Dark 20.
    Dark20,

    /// Pivot Style Dark 21.
    Dark21,

    /// Pivot Style Dark 22.
    Dark22,

    /// Pivot Style Dark 23.
    Dark23,

    /// Pivot Style Dark 24.
    Dark24,

    /// Pivot Style Dark 25.
    Dark25,

    /// Pivot Style Dark 26.
    Dark26,

    /// Pivot Style Dark 27.
    Dark27,

    /// Pivot Style Dark 28.
    Dark28,
}

impl fmt::Display for PivotTableStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, ""),
            Self::Light1 => write!(f, "PivotStyleLight1"),
            Self::Light2 => write!(f, "PivotStyleLight2"),
            Self::Light3 => write!(f, "PivotStyleLight3"),
            Self::Light4 => write!(f, "PivotStyleLight4"),
            Self::Light5 => write!(f, "PivotStyleLight5"),
            Self::Light6 => write!(f, "PivotStyleLight6"),
            Self::Light7 => write!(f, "PivotStyleLight7"),
            Self::Light8 => write!(f, "PivotStyleLight8"),
            Self::Light9 => write!(f, "PivotStyleLight9"),
            Self::Light10 => write!(f, "PivotStyleLight10"),
            Self::Light11 => write!(f, "PivotStyleLight11"),
            Self::Light12 => write!(f, "PivotStyleLight12"),
            Self::Light13 => write!(f, "PivotStyleLight13"),
            Self::Light14 => write!(f, "PivotStyleLight14"),
            Self::Light15 => write!(f, "PivotStyleLight15"),
            Self::Light16 => write!(f, "PivotStyleLight16"),
            Self::Light17 => write!(f, "PivotStyleLight17"),
            Self::Light18 => write!(f, "PivotStyleLight18"),
            Self::Light19 => write!(f, "PivotStyleLight19"),
            Self::Light20 => write!(f, "PivotStyleLight20"),
            Self::Light21 => write!(f, "PivotStyleLight21"),
            Self::Light22 => write!(f, "PivotStyleLight22"),
            Self::Light23 => write!(f, "PivotStyleLight23"),
            Self::Light24 => write!(f, "PivotStyleLight24"),
            Self::Light25 => write!(f, "PivotStyleLight25"),
            Self::Light26 => write!(f, "PivotStyleLight26"),
            Self::Light27 => write!(f, "PivotStyleLight27"),
            Self::Light28 => write!(f, "PivotStyleLight28"),
            Self::Medium1 => write!(f, "PivotStyleMedium1"),
            Self::Medium2 => write!(f, "PivotStyleMedium2"),
            Self::Medium3 => write!(f, "PivotStyleMedium3"),
            Self::Medium4 => write!(f, "PivotStyleMedium4"),
            Self::Medium5 => write!(f, "PivotStyleMedium5"),
            Self::Medium6 => write!(f, "PivotStyleMedium6"),
            Self::Medium7 => write!(f, "PivotStyleMedium7"),
            Self::Medium8 => write!(f, "PivotStyleMedium8"),
            Self::Medium9 => write!(f, "PivotStyleMedium9"),
            Self::Medium10 => write!(f, "PivotStyleMedium10"),
            Self::Medium11 => write!(f, "PivotStyleMedium11"),
            Self::Medium12 => write!(f, "PivotStyleMedium12"),
            Self::Medium13 => write!(f, "PivotStyleMedium13"),
            Self::Medium14 => write!(f, "PivotStyleMedium14"),
            Self::Medium15 => write!(f, "PivotStyleMedium15"),
            Self::Medium16 => write!(f, "PivotStyleMedium16"),
            Self::Medium17 => write!(f, "PivotStyleMedium17"),
            Self::Medium18 => write!(f, "PivotStyleMedium18"),
            Self::Medium19 => write!(f, "PivotStyleMedium19"),
            Self::Medium20 => write!(f, "PivotStyleMedium20"),
            Self::Medium21 => write!(f, "PivotStyleMedium21"),
            Self::Medium22 => write!(f, "PivotStyleMedium22"),
            Self::Medium23 => write!(f, "PivotStyleMedium23"),
            Self::Medium24 => write!(f, "PivotStyleMedium24"),
            Self::Medium25 => write!(f, "PivotStyleMedium25"),
            Self::Medium26 => write!(f, "PivotStyleMedium26"),
            Self::Medium27 => write!(f, "PivotStyleMedium27"),
            Self::Medium28 => write!(f, "PivotStyleMedium28"),
            Self::Dark1 => write!(f, "PivotStyleDark1"),
            Self::Dark2 => write!(f, "PivotStyleDark2"),
            Self::Dark3 => write!(f, "PivotStyleDark3"),
            Self::Dark4 => write!(f, "PivotStyleDark4"),
            Self::Dark5 => write!(f, "PivotStyleDark5"),
            Self::Dark6 => write!(f, "PivotStyleDark6"),
            Self::Dark7 => write!(f, "PivotStyleDark7"),
            Self::Dark8 => write!(f, "PivotStyleDark8"),
            Self::Dark9 => write!(f, "PivotStyleDark9"),
            Self::Dark10 => write!(f, "PivotStyleDark10"),
            Self::Dark11 => write!(f, "PivotStyleDark11"),
            Self::Dark12 => write!(f, "PivotStyleDark12"),
            Self::Dark13 => write!(f, "PivotStyleDark13"),
            Self::Dark14 => write!(f, "PivotStyleDark14"),
            Self::Dark15 => write!(f, "PivotStyleDark15"),
            Self::Dark16 => write!(f, "PivotStyleDark16"),
            Self::Dark17 => write!(f, "PivotStyleDark17"),
            Self::Dark18 => write!(f, "PivotStyleDark18"),
            Self::Dark19 => write!(f, "PivotStyleDark19"),
            Self::Dark20 => write!(f, "PivotStyleDark20"),
            Self::Dark21 => write!(f, "PivotStyleDark21"),
            Self::Dark22 => write!(f, "PivotStyleDark22"),
            Self::Dark23 => write!(f, "PivotStyleDark23"),
            Self::Dark24 => write!(f, "PivotStyleDark24"),
            Self::Dark25 => write!(f, "PivotStyleDark25"),
            Self::Dark26 => write!(f, "PivotStyleDark26"),
            Self::Dark27 => write!(f, "PivotStyleDark27"),
            Self::Dark28 => write!(f, "PivotStyleDark28"),
        }
    }
}

// -----------------------------------------------------------------------
// PivotTableFunction
// -----------------------------------------------------------------------

/// The `PivotTableFunction` enum defines the functions used to summarize
/// pivot table value fields.
///
/// These are the functions in the Excel "Summarize value field by" list. They
/// are used by the [`PivotTableValueField::set_function()`] method. The
/// default function is [`PivotTableFunction::Sum`].
///
/// [`PivotTableFunction::Count`] also counts text values, so it can be used
/// with non-numeric fields.
///
#[derive(Clone)]
pub enum PivotTableFunction {
    /// The sum of the values. This is the default.
    Sum,

    /// The number of values. This includes text values.
    Count,

    /// The average of the values.
    Average,

    /// The largest value.
    Max,

    /// The smallest value.
    Min,

    /// The product of the values.
    Product,

    /// The number of numeric values.
    CountNumbers,

    /// The standard deviation of a sample of the values.
    StdDev,

    /// The standard deviation of the population of the values.
    StdDevP,

    /// The variance of a sample of the values.
    Var,

    /// The variance of the population of the values.
    VarP,
}

impl PivotTableFunction {
    // The function name used in the default value field caption, such as
    // "Sum of Volume".
    fn caption(&self) -> &str {
        match self {
            Self::Sum => "Sum",
            Self::Max => "Max",
            Self::Min => "Min",
            Self::Var => "Var",
            Self::VarP => "Varp",
            Self::StdDev => "StdDev",
            Self::Average => "Average",
            Self::Product => "Product",
            Self::StdDevP => "StdDevp",
            Self::Count | Self::CountNumbers => "Count",
        }
    }
}

impl fmt::Display for PivotTableFunction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Sum => write!(f, "sum"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
            Self::Var => write!(f, "var"),
            Self::VarP => write!(f, "varp"),
            Self::Count => write!(f, "count"),
            Self::StdDev => write!(f, "stdDev"),
            Self::Average => write!(f, "average"),
            Self::Product => write!(f, "product"),
            Self::StdDevP => write!(f, "stdDevp"),
            Self::CountNumbers => write!(f, "countNums"),
        }
    }
}

// -----------------------------------------------------------------------
// Traits.
// -----------------------------------------------------------------------

/// Trait to map types into a pivot table data source range.
///
/// The [`PivotTable::set_data_source()`] method accepts the following types:
///
/// - A string with an Excel like range such as `"Sheet1!A1:D9"`. The range
///   can also use absolute `$` references and a quoted sheet name, like
///   `"'Sales Data'!$A$1:$D$9"`.
/// - A 5 value tuple that can be used to create the range programmatically
///   using a sheet name and zero indexed row and column values like:
///   `("Sheet1", 0, 0, 8, 3)` (this gives the same range as the first string
///   value above).
///
pub trait IntoPivotTableSource {
    /// Trait function to turn a type into a sheet name and a zero indexed
    /// `(first_row, first_col, last_row, last_col)` range.
    ///
    /// # Errors
    ///
    /// Returns the original range string as the error if it isn't a valid
    /// range.
    fn new_pivot_table_source(&self) -> Result<(String, RowNum, ColNum, RowNum, ColNum), String>;
}

impl IntoPivotTableSource for &str {
    fn new_pivot_table_source(&self) -> Result<(String, RowNum, ColNum, RowNum, ColNum), String> {
        parse_source_range(self).ok_or_else(|| self.to_string())
    }
}

impl IntoPivotTableSource for &String {
    fn new_pivot_table_source(&self) -> Result<(String, RowNum, ColNum, RowNum, ColNum), String> {
        parse_source_range(self).ok_or_else(|| (*self).clone())
    }
}

impl IntoPivotTableSource for (&str, RowNum, ColNum, RowNum, ColNum) {
    fn new_pivot_table_source(&self) -> Result<(String, RowNum, ColNum, RowNum, ColNum), String> {
        Ok((self.0.to_string(), self.1, self.2, self.3, self.4))
    }
}

// Parse a range string like "Sheet1!A1:D9" or "'Sheet 1'!$A$1:$D$9" into a
// sheet name and a zero indexed range.
fn parse_source_range(range_string: &str) -> Option<(String, RowNum, ColNum, RowNum, ColNum)> {
    let range_string = range_string.strip_prefix('=').unwrap_or(range_string);

    let (sheet_name, range) = utility::split_local_name(range_string)?;
    let sheet_name = utility::unquote_sheetname(sheet_name);
    if sheet_name.is_empty() {
        return None;
    }

    let range = range.replace('$', "").to_uppercase();
    if !utility::is_valid_range(&range) {
        return None;
    }

    let (first_cell, last_cell) = range.split_once(':')?;
    let (first_row, first_col) = parse_cell(first_cell)?;
    let (last_row, last_col) = parse_cell(last_cell)?;

    Some((sheet_name, first_row, first_col, last_row, last_col))
}

// Parse an "A1" style cell reference into zero indexed row and column values.
fn parse_cell(cell: &str) -> Option<(RowNum, ColNum)> {
    let (col_string, row_string) = utility::split_cell_reference(cell);

    // Limit the column name length to avoid an overflow. Columns beyond the
    // Excel limit are rejected in validate().
    if col_string.is_empty() || col_string.len() > 3 {
        return None;
    }

    let row = row_string.parse::<RowNum>().ok()?.checked_sub(1)?;
    let col = utility::column_name_to_number(&col_string);

    Some((row, col))
}

// -----------------------------------------------------------------------
// PivotTableSource
// -----------------------------------------------------------------------

// The worksheet range of the source data for a pivot table. Pivot tables with
// the same source share a pivot cache.
#[derive(Clone, Default, PartialEq)]
pub(crate) struct PivotTableSource {
    pub(crate) sheet_name: String,
    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) last_row: RowNum,
    pub(crate) last_col: ColNum,
    pub(crate) field_names: Vec<String>,
}

impl PivotTableSource {
    // The number of fields (columns) in the source data.
    pub(crate) fn num_fields(&self) -> ColNum {
        self.last_col - self.first_col + 1
    }

    // The number of data records (rows) in the source data, excluding the
    // header row.
    pub(crate) fn num_records(&self) -> RowNum {
        self.last_row - self.first_row
    }
}
