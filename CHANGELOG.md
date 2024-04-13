# Changelog

All notable changes to rust_xlsxwriter will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.64.2] - 2023-04-13

### Fixed

  - Fixed internal links in table of contents.


## [0.64.1] - 2023-03-26

### Added

  - Added the [`worksheet::set_screen_gridlines()`] method to turn on/offscreen gridlines.

  [`worksheet::set_screen_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_screen_gridlines

  - Added updated docs on [Working with Workbooks] and [Working with Worksheets].

  [Working with Workbooks]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/index.html
  [Working with Worksheets]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/index.html


## [0.64.0] - 2023-03-18

### Added

- Add support for worksheet sparklines. Sparklines are a feature of Excel 2010+
  which allows you to add small charts to worksheet cells. These are useful for
  showing data trends in a compact visual format.

  See [Working with Sparklines].

  [Working with Sparklines]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/sparkline/index.html


## [0.63.0] - 2023-02-25

### Added

- Added support for embedding images into worksheets with
  [`worksheet::embed_image()`] and [`worksheet::embed_image_with_format()`] and
  the [Image] struct. See the [Embedded Images] example.

  This can be useful if you are building up a spreadsheet of products with a
  column of images for each product. Embedded images move with the cell so they
  can be used in worksheet tables or data ranges that will be sorted or
  filtered.

  This functionality is the equivalent of Excel's menu option to insert an image
  using the option to "Place in Cell" which is available in Excel 365 versions
  from 2023 onwards.

  [`worksheet::embed_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.embed_image
  [`worksheet::embed_image_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.embed_image_with_format

  [Embedded Images]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html#insert-images-embedding-an-image-in-a-cell


- Updated polars dependency to 0.37.2 to pick up latest Polars additions for
  [`polars_excel_writer`].


- Added [`utility::check_sheet_name()`] function to allow checking for valid
  worksheet names according to Excel's naming rules. This functionality was
  previously `pub(crate)` private.

  [Feature Request #83].

  [Feature Request #83]: https://github.com/jmcnamara/rust_xlsxwriter/pull/83

  [`utility::check_sheet_name()`]:
      https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.check_sheet_name.html

### Removed

- Removed unnecessary lifetime on [`Format`] objects used in Worksheet `write()`
  methods. This allows the the [`IntoExcelData`] trait to be defined for user
  types and have them include a default format. See [Feature Request #85].

  [Feature Request #85]: https://github.com/jmcnamara/rust_xlsxwriter/pull/85
  [`Format`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html

## [0.62.0] - 2023-01-24

### Added

- Added support for adding a worksheet [`Table`] as a serialization format. See [`SerializeFieldOptions::set_table()`].

- Added [`Worksheet::get_serialize_dimensions()`] and
  [`Worksheet::get_serialize_column_dimensions()`] methods to get dimensions
  from a serialized range.

- Updated polars dependency to 0.36.2 to pick up Polars `AnyData` changes for
  [`polars_excel_writer`].


[`Worksheet::get_serialize_dimensions()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.get_serialize_dimensions

[`Worksheet::get_serialize_column_dimensions()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.get_serialize_column_dimensions


[`SerializeFieldOptions::set_table()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/struct.SerializeFieldOptions.html#method.set_table


### Changed

- Changed APIs for [`Table`] to return `Table` instead of `&Table` to allow
  methods to be chained. This makes worksheet Table usage easier during
  serialization. Note that this is a backward incompatible change.



## [0.61.0] - 2023-01-13

### Added

- Added support for a `XlsxSerialize` derive and struct attributes to control
  the formatting and options of the Excel output during serialization. These are
  similar in intention to Serde container/field attributes.

  See [Controlling Excel output via `XlsxSerialize` and struct attributes] and
  [Working with Serde].

  [Feature Request #66].

  [Feature Request #66]: https://github.com/jmcnamara/rust_xlsxwriter/pull/66

  [Controlling Excel output via `XlsxSerialize` and struct attributes]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html#controlling-excel-output-via-xlsxserialize-and-struct-attributes


- Added `XlsxError::CustomError` as a target error for external crates/apps.

  [Feature Request #72].

  [Feature Request #72]: https://github.com/jmcnamara/rust_xlsxwriter/pull/72


## [0.60.0] - 2023-01-02

### Added

- Added support for setting Serde headers using deserialization of a target
  struct type as well as the previous method of using serialization and an
  instance of the struct type. See [Setting serialization headers].

  [Feature Request #63].

- Added additional support for serialization header and field options via
  [`CustomSerializeField`].

- Added support for writing `Result<T, E>` with [`worksheet.write()`] when `T`
  and `E` are supported types.

  [Feature Request #64].

[Feature Request #63]: https://github.com/jmcnamara/rust_xlsxwriter/pull/63
[Feature Request #64]: https://github.com/jmcnamara/rust_xlsxwriter/pull/64
[Setting serialization headers]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html#setting-serialization-headers
[`CustomSerializeField`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/struct.CustomSerializeField.html

## [0.59.0] - 2023-12-15

### Added

- Added [`serialize_chrono_option_naive_to_excel()`] to help serialization of
  `Option` Chrono types. [Feature Request #62].

[Feature Request #62]: https://github.com/jmcnamara/rust_xlsxwriter/pull/62

[`serialize_chrono_option_naive_to_excel()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.serialize_chrono_option_naive_to_excel.html


## [0.58.0] - 2023-12-11

### Added

- Added serialization support for [`ExcelDateTime`] and [`Chrono`] date/time
  types. See [Working with Serde - Serializing dates and times].

  [Working with Serde - Serializing dates and times]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html#serializing-dates-and-times


## [0.57.0] - 2023-12-09

### Added

- Added support for Serde serialization. This requires the `serde` feature flag
  to be enabled. See [Working with Serde].


- Added support for writing `u64` and `i64` number within Excel's limitations.
  This implies a loss of precision outside Excel's integer range of +/-
  999,999,999,999,999 (15 digits).

  [Working with Serde]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html#working-with-serde

## [0.56.0] - 2023-11-27

### Added

- Changed some of the Conditional Format interfaces introduced in the previous
  release to use extended enums. This is an API change with the version released
  earlier this week but it provides a cleaner interface.

- Added support for `Option<T>` wrapped types to [`worksheet.write()`].

  [Feature Request #59].

[Feature Request #59]: https://github.com/jmcnamara/rust_xlsxwriter/pull/59


## [0.55.0] - 2023-11-21

### Added

- Added support for conditional formatting. See [Working with Conditional Formats].

[Working with Conditional Formats]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/conditional_format/index.html


## [0.54.0] - 2023-11-04

### Added

- Added option to add a chart data table to charts via the
  [`Chart::set_data_table()`] and [`ChartDataTable`].

- Added option to set the display units on a Y-axis to units such as Thousands or
  Millions via the [`ChartAxis::set_display_unit_type()`] method.

- Added option to set the crossing position of axes via the [`ChartAxis::set_crossing()`] method.

- Added option to set the axes label alignment via the [`ChartAxis::set_label_alignment()`] method.

- Added option to turn on/off line smoothing for Line and Scatter charts via the
  [`ChartSeries::set_smooth()`] method.

[`ChartDataTable`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartDataTable.html
[`Chart::set_data_table()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.set_data_table
[`ChartAxis::set_crossing()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_crossing
[`ChartAxis::set_label_alignment()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_label_alignment
[`ChartAxis::set_display_unit_type()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_display_unit_type
[`ChartSeries::set_smooth()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_smooth


## [0.53.0] - 2023-10-30

### Added

- Added support for Excel Stock charts. See the [Stock Chart] cookbook example.

- Added support to charts for:
  - Up-Down bars via the [`Chart::set_up_down_bars()`] struct and methods.
  - High-Low lines via the [`Chart::set_high_low_lines()`] struct and methods.
  - Drop lines via the [`Chart::set_high_low_lines()`] struct and methods.
  - Chart axis support for Date, Text and Automatic axes via the
    [`ChartAxis::set_date_axis()`], [`ChartAxis::set_text_axis()`]
    and[`ChartAxis::set_automatic_axis()`] methods.
  - Chart axis support for minimum and maximum date values via the
    [`ChartAxis::set_min_date()`] and [`ChartAxis::set_max_date()`] methods.

- Add worksheet syntactic helper methods
  [`Worksheet::write_row_with_format()`] and
  [`Worksheet::write_column_with_format()`].

[`Chart::set_drop_lines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.set_drop_lines
[`Chart::set_up_down_bars()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.set_up_down_bars
[`Chart::set_high_low_lines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.set_high_low_lines

[`ChartAxis::set_date_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_date_axis
[`ChartAxis::set_text_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_text_axis
[`ChartAxis::set_automatic_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_automatic_axis

[`ChartAxis::set_min_date()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_min_date
[`ChartAxis::set_max_date()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_max_date

[`Worksheet::write_row_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_row_with_format
[`Worksheet::write_column_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_column_with_format


[Stock Chart]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html#chart-stock-excel-stock-chart-example

## [0.52.0] - 2023-10-20

### Added

- Added support for chart series error bars via the [`ChartErrorBars`] struct
  and methods.

[`ChartErrorBars`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartErrorBars.html

### Fixed

- Fixed XML error in non-Pie charts.

  [GitHub Issue #55].

[GitHub Issue #55]: https://github.com/jmcnamara/rust_xlsxwriter/issues/55



## [0.51.0] - 2023-10-15

### Added

- Added support for chart gradient fill formatting via the [`ChartGradientFill`] struct and methods.

- Added support for formatting the chart trendlines data labels via the
  [`ChartTrendline::set_label_font`] and [`ChartTrendline::set_label_format`].


[`ChartGradientFill`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartGradientFill.html
[`ChartTrendline::set_label_font`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartTrendline.html#method.set_label_font
[`ChartTrendline::set_label_format`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartTrendline.html#method.set_label_format


## [0.50.0] - 2023-10-12

### Added

- Added support for chart trendlines (Linear, Polynomial, Moving Average, etc.)
  via the [`ChartTrendline`] struct and methods.

- Added the [`worksheet.set_very_hidden()`] method to hide a worksheet similar
  to the [`worksheet.set_hidden()`] method. The difference is that the worksheet
  can only be unhidden by VBA and cannot be unhidden in the the Excel user
  interface.

- Added support for leader lines to non-Pie charts.

### Fixed

- Fixed handling of [future functions] in table formulas.


[`ChartTrendline`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartTrendline.html
[`worksheet.set_very_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_very_hidden
[future functions]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Formula.html#formulas-added-in-excel-2010-and-later


## [0.49.0] - 2023-09-19

### Added

- Added chart options to control how non-data cells are displayed.

  - [`chart.show_empty_cells_as()`]
  - [`chart.show_na_as_empty_cell()`]
  - [`chart.show_hidden_data()`]

[`chart.show_hidden_data()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.show_hidden_data
[`chart.show_empty_cells_as()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.show_empty_cells_as
[`chart.show_na_as_empty_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.show_na_as_empty_cell


- Updated Polar's dependency and `PolarError` import to reflect changes in Polars v 0.33.2.

## [0.48.0] - 2023-09-08

### Added

- Added support for custom total formulas to [`TableFunction`].

[`TableFunction`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.TableFunction.html


## [0.47.0] - 2023-09-02

### Added

- Added `wasm` feature flag to help compilation on Wasm/Javascript targets. Also
  added mapping from a `XlsxError` to a `JsValue` error.

  See the [rust_xlsx_wasm_example] sample application that demonstrates
  accessing `rust_xlsxwriter` code from JavaScript, Node.js, Deno and Wasmtime.

- Added [`workbook.save_to_writer()`] method to make it easier to interact with
  interfaces that implement the `<W: Write>` trait.

[rust_xlsx_wasm_example]: https://github.com/Clipi-12/rust_xlsx_wasm_example
[`workbook.save_to_writer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.save_to_writer


## [0.46.0] - 2023-08-20

### Added

- Added `polars` feature flag to help interoperability with Polars. Currently
  it only implements `PolarsError` and `XlsxError` mapping but other
  functionality may be added in the future. These changes are added to support
  the [`polars_excel_writer`] crate.

  [`polars_excel_writer`]: https://crates.io/crates/polars_excel_writer


## [0.45.0] - 2023-08-12

### Fixed

- Fixed "multiply with overflow" issue when image locations in the worksheet
  were greater than the maximum `u32` value.

  Related to [GitHub Issue #51].

[GitHub Issue #51]: https://github.com/jmcnamara/rust_xlsxwriter/issues/51


## [0.44.0] - 2023-08-02

### Added

- Added threading into the backend worksheet writing for increased performance
  with large multi-worksheet files.


## [0.43.0] - 2023-07-27

### Added

- Added support for worksheet [`Table`] header and cell formatting via the
  [`tablecolumn.set_format()`] and [`tablecolumn.set_header_format()`] methods.

[`tablecolumn.set_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.TableColumn.html#method.set_format
[`tablecolumn.set_header_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.TableColumn.html#method.set_header_format


## [0.42.0] - 2023-07-11

### Changed

- Made the `chrono` feature optionally off instead of optionally on. The
  `chrono` feature must now be explicitly enabled to allow support for
  [`Chrono`] types.

- Renamed the worksheet `write_datetime()` method to the API consistent
  [`write_datetime_with_format()`] and introduced a new
  [`write_datetime()`] method that doesn't take a format. This is
  required to fix a error in the APIs that prevented an unformatted datetime
  from taking the row or column format.

  **Note**: This is a backwards incompatible change.

  See [GitHub Issue #47].


### Added

- Added a [Tutorial] and [Cookbook] section to the `doc.rs` documentation.

- Added a check, and and error result, for case-insensitive duplicate sheet
  names. Also added sheet name validation to chart series.

  See [GitHub Issue #45].

- Added cell range name handling utility functions:

  - [`column_number_to_name()`] - Convert a zero indexed column cell reference
    to a string like `"A"`.
  - [`column_name_to_number()`] - Convert a column string such as `"A"` to a
    zero indexed column reference.
  - [`row_col_to_cell()`] - Convert zero indexed row and column cell numbers to
    a `A1` style string.
  - [`row_col_to_cell_absolute()`] - Convert zero indexed row and column cell
    numbers to an absolute `$A$1` style range string.
  - [`cell_range()`] - Convert zero indexed row and col cell numbers to a
    `A1:B1` style range string.
  - [`cell_range_absolute()`] - Convert zero indexed row and col cell numbers to
    an absolute `$A$1:$B$1`

[`cell_range()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/fn.cell_range.html
[`row_col_to_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/fn.row_col_to_cell.html
[`cell_range_absolute()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/fn.cell_range_absolute.html
[`column_number_to_name()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/fn.column_number_to_name.html
[`column_name_to_number()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/fn.column_name_to_number.html
[`row_col_to_cell_absolute()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/fn.row_col_to_cell_absolute.html

[GitHub Issue #45]: https://github.com/jmcnamara/rust_xlsxwriter/issues/45
[GitHub Issue #47]: https://github.com/jmcnamara/rust_xlsxwriter/issues/47

[Tutorial]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/tutorial/index.html
[Cookbook]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html
[`write_datetime()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_datetime
[`write_datetime_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_datetime_with_format


## [0.41.0] - 2023-06-20

- Added the native [`ExcelDateTime`] struct to allow handling of dates and times
  without a dependency on the [`Chrono`] library. The Chrono library is now an
  optional feature/dependency. It is included by default in this release for
  compatibility with previous versions but it will be optionally off in the next
  and subsequent versions.

  All date/time APIs support both the native `ExcelDateTime` and `Chrono` types
  via the [`IntoExcelDateTime`] trait.

  The `worksheet.write_date()` and `worksheet.write_time()` methods have been
  moved to "undocumented" since the same functionality is available via
  [`worksheet.write_datetime()`]. This is a soft deprecation.

[`Chrono`]: https://docs.rs/chrono/latest/chrono
[`ExcelDateTime`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html
[`IntoExcelDateTime`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/trait.IntoExcelDateTime.html
[`worksheet.write_datetime()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_datetime


## [0.40.0] - 2023-05-31

- Added support for worksheet tables. See [`Table`] and the
  [`worksheet.add_table()`] method.

- Added support for the `bool` type in the generic [`worksheet.write()`] method.

[`Table`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Table.html
[`worksheet.add_table()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.add_table


## [0.39.0] - 2023-05-23

### Added

- Added [`worksheet.write_row()`], [`worksheet.write_column()`],
  [`worksheet.write_row_matrix()`] and [`worksheet.write_column_matrix()`]
  methods to write arrays/iterators of data.

- Added [`Formula`] and [`Url`] types to use with generic [`worksheet.write()`].

  [Feature Request #16].

- Make several string handling APIs more generic using `impl Into<String>`.

  [Feature Request #16].

- Renamed/refactored `XlsxColor` to [`Color`] for API consistency. The
  `XlsxColor` type alias is still available for backward compatibility.

[`Url`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Url.html
[`Color`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.Color.html
[`Formula`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Formula.html
[Feature Request #16]: https://github.com/jmcnamara/rust_xlsxwriter/pull/16
[`worksheet.write()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write
[`worksheet.write_row()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_row
[`worksheet.write_column()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_column
[`worksheet.write_row_matrix()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_row_matrix
[`worksheet.write_column_matrix()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_column_matrix


## [0.38.0] - 2023-05-05

### Added

- Added several Chart axis and series methods:

  - [`chartaxis.set_hidden()`]
  - [`chartaxis.set_label_interval()`]
  - [`chartaxis.set_label_position()`]
  - [`chartaxis.set_log_base()`]
  - [`chartaxis.set_major_gridlines()`]
  - [`chartaxis.set_major_gridlines_line()`]
  - [`chartaxis.set_major_tick_type()`]
  - [`chartaxis.set_major_unit()`]
  - [`chartaxis.set_max()`]
  - [`chartaxis.set_min()`]
  - [`chartaxis.set_minor_gridlines()`]
  - [`chartaxis.set_minor_gridlines_line()`]
  - [`chartaxis.set_minor_tick_type()`]
  - [`chartaxis.set_minor_unit()`]
  - [`chartaxis.set_position_between_ticks()`]
  - [`chartaxis.set_reverse()`]
  - [`chartaxis.set_tick_interval()`]
  - [`chartseries.set_invert_if_negative()`]
  - [`chartseries.set_invert_if_negative_color()`]

[`chartaxis.set_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_hidden
[`chartaxis.set_label_interval()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_label_interval
[`chartaxis.set_label_position()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_label_position
[`chartaxis.set_log_base()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_log_base
[`chartaxis.set_major_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_major_gridlines
[`chartaxis.set_major_gridlines_line()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_major_gridlines_line
[`chartaxis.set_major_tick_type()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_major_tick_type
[`chartaxis.set_major_unit()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_major_unit
[`chartaxis.set_max()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_max
[`chartaxis.set_min()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_min
[`chartaxis.set_minor_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_minor_gridlines
[`chartaxis.set_minor_gridlines_line()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_minor_gridlines_line
[`chartaxis.set_minor_tick_type()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_minor_tick_type
[`chartaxis.set_minor_unit()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_minor_unit
[`chartaxis.set_position_between_ticks()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_position_between_ticks
[`chartaxis.set_reverse()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_reverse
[`chartaxis.set_tick_interval()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartAxis.html#method.set_tick_interval
[`chartseries.set_invert_if_negative()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_invert_if_negative
[`chartseries.set_invert_if_negative_color()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_invert_if_negative_color


## [0.37.0] - 2023-04-30

### Added

- Added font formatting support to chart titles, legends, axes and data labels
  via [`ChartFont`] and various `set_font()` methods.

- Made [`worksheet.write_string()`] and [`worksheet.write()`] more generic via
  `impl Into<String>` to allow them to handle `&str`, `&String`, `String`, and
  `Cow<>` types.

  See [GitHub Feature Request #35].

[`ChartFont`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartFont.html
[`worksheet.write()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_string
[`worksheet.write_string()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_string
[GitHub Feature Request #35]: https://github.com/jmcnamara/rust_xlsxwriter/issues/35


## [0.36.1] - 2023-04-18

Fix cargo/release issue with 0.36.0 release.

## [0.36.0] - 2023-04-18

### Added

- Added performance improvement for applications that use a lot of `Format`
  objects. [GitHub Issue #30].


### Fixed

- Fixed issue introduced in v0.34.0 where `Rc<>` value was blocking `Send` in
  multithreaded applications. [GitHub Issue #29].

[GitHub Issue #29]: https://github.com/jmcnamara/rust_xlsxwriter/issues/29
[GitHub Issue #30]: https://github.com/jmcnamara/rust_xlsxwriter/issues/30


## [0.35.0] - 2023-04-16

### Added

- Added support for Chart Series data labels including custom data labels. See
  [`ChartDataLabel`], [`chart.series.set_data_label()`] and [`chart.series.set_custom_data_labels()`].

[`ChartDataLabel`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartDataLabel.html
[`chart.series.set_data_label()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_data_label
[`chart.series.set_custom_data_labels()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_custom_data_labels


## [0.34.0] - 2023-04-12

### Added

Performance improvement release.

- Added optimizations across the library. For larger files this version is 10%
  faster than previous versions.

  These optimizations were provided by Adrián Delgado, see [GitHub Issue #23].

- Added crate feature `zlib` which adds a dependency on zlib and a C compiler
  but is around 1.6x faster for larger files. With this feature enabled it is
  even faster than the native C version libxlsxwriter by around 1.4x for large
  files.

  See also the [Performance] section of the user guide.

[GitHub Issue #23]: https://github.com/jmcnamara/rust_xlsxwriter/issues/23
[Performance]: https://rustxlsxwriter.github.io/performance.html

## [0.33.0] - 2023-04-10

### Added

- Added support for formatting and setting chart points via the [`ChartPoint`]
  struct. This is mainly useful as the way of specifying segment colors in Pie
  charts.

  See the updated [Pie Chart] example in the user guide.

- Added support for formatting and setting chart markers via the [`ChartMarker`]
  struct.

- Added [`chart.set_rotation()`] and [`chart.set_hole_size()`] methods for Pie and Doughnut charts.

- Added support to differentiate between `XlsxColor::Default` and
  `XlsxColor::Automatic` colors for Excel elements. These are usually equivalent
  but there are some cases where the "Automatic" color, which can be set at a
  system level, is different from the Default color.

[Pie Chart]: https://rustxlsxwriter.github.io/examples/pie_chart.html
[`ChartPoint`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartPoint.html
[`ChartMarker`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartMarker.html
[`chart.set_rotation()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.set_rotation
[`chart.set_hole_size()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html#method.set_hole_size


## [0.32.0] - 2023-04-03

### Added

- Added formatting for the chart title and axes via the the [`ChartFormat`]
  struct.

## [0.31.0] - 2023-04-02

### Added

- Added formatting for the chart area, plot area, and legend via the the
  [`ChartFormat`] struct.


## [0.30.0] - 2023-03-31

### Added

- Added chart formatting for Lines, Borders, Solid fills and Pattern fills via
  the [`ChartFormat`] struct. This is currently only available for chart series
  but it will be extended in the next release for most other chart elements.

  See also the [Chart Fill Pattern] example in the user guide.

- Added [`IntoColor`] trait to allow syntactic shortcuts for [`XlsxColor`]
  parameters in methods. So now you can set a RGB color like this
  `object.set_color("#FF7F50")` instead of the more verbose
  `object.set_color(XlsxColor::RGB(0xFF7F50))`. This addition doesn't require
  any API changes from the end user.

- Added [`worksheet.insert_image_fit_to_cell()`] method to add an image to a
  worksheet and scale it so that it fits in a cell. This method can be useful
  when creating a product spreadsheet with a column of images for each product.

  See also the [insert_image_to_fit] example in the user guide.

- Added [`chart.series.set_gap()`] and [`chart.series.set_overlap()`] method to control layout
  of histogram style charts.

[`IntoColor`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/trait.IntoColor.html
[`ChartFormat`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartFormat.html
[Chart Fill Pattern]: https://rustxlsxwriter.github.io/examples/chart_pattern.html
[insert_image_to_fit]: https://rustxlsxwriter.github.io/examples/insert_image_to_fit.html
[`chart.series.set_gap()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_gap
[`chart.series.set_overlap()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartSeries.html#method.set_overlap
[`worksheet.insert_image_fit_to_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.insert_image_fit_to_cell


## [0.29.0] - 2023-03-16

### Added

- Added support for resizing and object positioning to the [`Chart`] struct.

- Added handling for `chrono` date/time types to the generic
  [`worksheet.write()`] method.


## [0.28.0] - 2023-03-14

### Added

- Added support for positioning or hiding Chart legends. See [`ChartLegend`].

[`ChartLegend`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ChartLegend.html


## [0.27.0] - 2023-03-13

### Added

- Added support for Charts via the [`Chart`] struct and the
  [`worksheet.insert_chart()`] method. See also the [Chart Examples] in the user
  guide.

- Added a generic [`worksheet.write()`] method that writes string or number
  types. This will be extended in an upcoming release to provide a single
  `write()` method for all of the currently supported types.

  It also allows the user to extend [`worksheet.write()`] to handle user defined
  types via the [`IntoExcelData`] trait. See also the [Writing Generic data]
  example in the user guide.

[`Chart`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Chart.html
[Chart Examples]: https://rustxlsxwriter.github.io/examples/simple_chart.html
[`IntoExcelData`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/trait.IntoExcelData.html
[`worksheet.write()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write
[Writing Generic data]: https://rustxlsxwriter.github.io/examples/generic_write.html
[`worksheet.insert_chart()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.insert_chart


## [0.26.0] - 2023-02-03

  Note: this version contains a major refactoring/renaming of some of the main
  data writing functions and some of the enums and secondary structs. This will
  require code changes from all current users but will allow more consistent
  APIs in future releases. Nevertheless, I apologize for this level of change.


### Changed

- The following worksheet functions have changed names to reflect their
  frequency of usage.

  | Previous name                        | New name                                    |
  | :----------------------------------- | :------------------------------------------ |
  | `write_string_only()`                | `write_string()`                            |
  | `write_number_only()`                | `write_number()`                            |
  | `write_formula_only()`               | `write_formula()`                           |
  | `write_boolean_only()`               | `write_boolean()`                           |
  | `write_rich_string_only()`           | `write_rich_string()`                       |
  | `write_array_formula_only()`         | `write_array_formula()`                     |
  | `write_dynamic_array_formula_only()` | `write_dynamic_array_formula()`             |
  |                                      |                                             |
  | `write_array_formula()`              | `write_array_formula_with_format()`         |
  | `write_boolean()`                    | `write_boolean_with_format()`               |
  | `write_dynamic_array_formula()`      | `write_dynamic_array_formula_with_format()` |
  | `write_formula()`                    | `write_formula_with_format()`               |
  | `write_number()`                     | `write_number_with_format()`                |
  | `write_rich_string()`                | `write_rich_string_with_format()`           |
  | `write_string()`                     | `write_string_with_format()`                |


- The following enums and structs have changed to a more logical naming:

  | Previous name             | New name               |
  | :------------------------ | :--------------------- |
  | `XlsxAlign`               | `FormatAlign`          |
  | `XlsxBorder`              | `FormatBorder`         |
  | `XlsxDiagonalBorder`      | `FormatDiagonalBorder` |
  | `XlsxPattern`             | `FormatPattern`        |
  | `XlsxScript`              | `FormatScript`         |
  | `XlsxUnderline`           | `FormatUnderline`      |
  |                           |                        |
  | `XlsxObjectMovement`      | `ObjectMovement`       |
  | `XlsxImagePosition`       | `HeaderImagePosition`  |
  |                           |                        |
  | `ProtectWorksheetOptions` | `ProtectionOptions`    |
  | `Properties`              | `DocProperties`        |


- The `DocProperties::set_custom_property()` method replaces several type
  specific methods with a single trait based generic method.


## [0.25.0] - 2023-01-30

### Added

- Added ability to filter columns in [`worksheet.autofilter()`] ranges via
  [`worksheet.filter_column()`] and [`FilterCondition`].

  The library automatically hides any rows that don't match the supplied
  criteria. This is an additional feature that isn't available in the other
  language ports of "xlsxwriter".

  See also the [Working with Autofilters] section of the Users Guide.

[`FilterCondition`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.FilterCondition.html
[Working with Autofilters]: https://rustxlsxwriter.github.io/formulas/autofilters.html
[`worksheet.autofilter()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.autofilter
[`worksheet.filter_column()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.filter_column


## [0.24.0] - 2023-01-18

### Added

- Added support for hiding rows and columns (to hide intermediate calculations)
  via the [`worksheet.set_column_hidden()`] and[`worksheet.set_row_hidden()`]
  method. This is also a required precursor to adding autofilter conditions.
- Added the [ObjectMovement] enum to control how a worksheet object, such a
  an image, moves when the cells underneath it are moved, resized or deleted.

[ObjectMovement]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.ObjectMovement.html
[`worksheet.set_row_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_row_hidden
[`worksheet.set_column_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_column_hidden

## [0.23.0] - 2023-01-16

### Added

Added more page setup methods.

- Added [`worksheet.set_selection()`] method to select a cell or range of cells in a worksheet.
- Added [`worksheet.set_top_left_cell()`] method to set the top and leftmost visible cell.
- Added [`worksheet.set_page_breaks()`] method to add page breaks to a worksheet.

[`worksheet.set_selection()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_selection
[`worksheet.set_page_breaks()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_page_breaks
[`worksheet.set_top_left_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_top_left_cell


## [0.22.0] - 2023-01-13

### Added

- Added support for worksheet protection via the [`worksheet.protect()`],
  [`worksheet.protect_with_password()`] and [`worksheet.protect_with_options()`].

  See also the section on [Worksheet protection] in the user guide.

- Add option to make the xlsx file read-only when opened by Excel via the
  [`workbook.read_only_recommended()`] method.


[Worksheet protection]:  https://rustxlsxwriter.github.io/worksheet/protection.html
[`worksheet.protect()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.protect
[`worksheet.protect_with_options()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.protect_with_options
[`workbook.read_only_recommended()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.read_only_recommended
[`worksheet.protect_with_password()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.protect_with_password


## [0.21.0] - 2023-01-09

### Added

- Added support for setting document metadata properties such as Author and
  Creation Date. For more details see [`DocProperties`] and
  [`workbook::set_properties()`].

[`DocProperties`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.DocProperties.html
[`workbook::set_properties()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.set_properties

### Changed

- Change date/time parameters to references in [`worksheet.write_datetime()`],
  [`worksheet.write_date()`] and [`worksheet.write_time()`] for consistency.

[`worksheet.write_date()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_date
[`worksheet.write_time()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_time
[`worksheet.write_datetime()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_datetime


## [0.20.0] - 2023-01-06

### Added

- Improved fitting algorithm for [`worksheet.autofit()`]. See also the
  [app_autofit] sample application.

### Changed

- The `worksheet.set_autofit()` method has been renamed to `worksheet.autofit()`
  for consistency with the other language versions of this library.


## [0.19.0] - 2022-12-27

### Added

- Added support for created defined variable names at a workbook and worksheet
  level via [`workbook.define_name()`].

  See also [Using defined names] in the user guide.

- Added initial support for autofilters via [`worksheet.autofilter()`].

  Note, adding filter criteria isn't currently supported. That will be added in
  an upcoming version. See also [Adding Autofilters] in the user guide.

[Adding Autofilters]: https://rustxlsxwriter.github.io/examples/autofilter.html
[Using defined names]: https://rustxlsxwriter.github.io/examples/defined_names.html
[`worksheet.autofilter()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.autofilter
[`workbook.define_name()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.define_name


## [0.18.0] - 2022-12-19

### Added

- Added support for "rich" strings with multiple font formats via
  [`worksheet.write_rich_string()`] and [`worksheet.write_rich_string_with_format()`].
  For example strings like "This is **bold** and this is *italic*".

  See also the [Rich strings example] in the user guide.

[Rich strings example]: https://rustxlsxwriter.github.io/examples/rich_strings.html
[`worksheet.write_rich_string_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_rich_string
[`worksheet.write_rich_string()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_rich_string_only

## [0.17.1] - 2022-12-18

### Fixed

- Fixes issue where header image files became corrupt during incremental saves.
  Also fixes similar issues in some formatting code.


## [0.17.0] - 2022-12-17

### Added

- Added support for images in headers/footers via the
  [`worksheet.set_header_image()`] and [`worksheet.set_footer_image()`] methods.

  See the [Headers and Footers] and [Adding a watermark] examples in the user guide.

[Headers and Footers]: https://rustxlsxwriter.github.io/examples/headers.html
[Adding a watermark]: https://rustxlsxwriter.github.io/examples/watermark.html
[`worksheet.set_footer_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_footer_image
[`worksheet.set_header_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_header_image


## [0.16.0] - 2022-12-09

### Added

- Replicate the optimization used by Excel where it only stores one copy of a
  repeated/duplicate image in a workbook.


## [0.15.0] - 2022-12-08

### Added

- Added support for images in buffers via [`image.new_from_buffer()`].

- Added image accessability features via [`image.set_alt_text()`] and[`image.set_decorative()`].

[`image.set_alt_text()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Image.html#method.set_alt_text
[`image.set_decorative()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Image.html#method.set_decorative
[`image.new_from_buffer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Image.html#method.new_from_buffer


## [0.14.0] - 2022-12-05

### Added

- Added support for inserting images into worksheets with
  [`worksheet.insert_image()`] and [`worksheet.insert_image_with_offset()`] and
  the [Image] struct.

  See also the [images example] in the user guide.

  Upcoming versions of the library will support additional image handling
  features such as EMF and WMF formats, removal of duplicate images, hyperlinks
  in images and images in headers/footers.

### Removed

- The [`workbook.save()`] method has been extended to handle paths or strings.
  The `workbook.save_to_path()` method has been removed. See [PR #15].

[Image]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html
[PR #15]: https://github.com/jmcnamara/rust_xlsxwriter/pull/15
[images example]: https://rustxlsxwriter.github.io/examples/images.html
[`worksheet.insert_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.insert_image
[`worksheet.insert_image_with_offset()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.insert_image_with_offset


## [0.13.0] - 2022-11-21

### Added

- Added support for writing hyperlinks in worksheets via the following methods:

  - [`worksheet.write_url()`] to write a link with a default hyperlink format style.
  - [`worksheet.write_url_with_text()`] to add alternative text to the link.
  - [`worksheet.write_url_with_format()`] to add an alternative format to the link.
  - [`worksheet.write_url_with_options()`] to add a screen tip and all other options to the link.

See also the [hyperlinks example] in the user guide.

[hyperlinks example]: https://rustxlsxwriter.github.io/examples/hyperlinks.html
[`worksheet.write_url()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_url
[`worksheet.write_url_with_text()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_url_with_text
[`worksheet.write_url_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_url_with_format
[`worksheet.write_url_with_options()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_url_with_options


## [0.12.1] - 2022-11-09

### Changed

- Dependency changes to make WASM compilation easier:

  - Reduced the `zip` dependency to the minimum import only.
  - Removed dependency on `tempfile`. The library now uses in memory files.

## [0.12.0] - 2022-11-06

### Added

- Added [`worksheet.merge_range()`] method.
- Added support for Theme colors to [`XlsxColor`]. See also [Working with
  Colors] in the user guide.

[`XlsxColor`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.XlsxColor.html
[Working with Colors]: https://rustxlsxwriter.github.io/colors/intro.html
[`worksheet.merge_range()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.merge_range


## [0.11.0] - 2022-11-04

### Added

- Added several worksheet methods for working with worksheet tabs:

  - [`worksheet.set_active()`]: Set the active/visible worksheet.
  - [`worksheet.set_tab_color()`]: Set the tab color.
  - [`worksheet.set_hidden()`]: Hide a worksheet.
  - [`worksheet.set_selected()`]: Set a worksheet as selected.
  - [`worksheet.set_first_tab()`]: Set the first visible tab.

  See also [Working with worksheet tabs] in the user guide.

[`worksheet.set_active()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_active
[`worksheet.set_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_hidden
[`worksheet.set_selected()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_selected
[`worksheet.set_tab_color()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_tab_color
[`worksheet.set_first_tab()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_first_tab
[Working with worksheet tabs]: https://rustxlsxwriter.github.io/worksheet/tabs.html


## [0.10.0] - 2022-11-03

### Added

- Added a simulated [`worksheet.autofit()`] method to automatically adjust
  the width of columns with data. See also the [app_autofit] sample application.

- Added the [`worksheet.set_freeze_panes()`] method to set "freeze" panes for
  worksheets. See also the [app_panes] example application.

[app_panes]: https://rustxlsxwriter.github.io/examples/panes.html
[app_autofit]: https://rustxlsxwriter.github.io/examples/autofit.html
[`worksheet.autofit()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.autofit
[`worksheet.set_freeze_panes()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_freeze_panes

## [0.9.0] - 2022-10-31

**Note, this version contains a major backward incompatible API change where it
restructures the Workbook constructor/destructor sequence and introduces a
`save()` method to replace `close()`.**

### Changed

- The [`Workbook::new()`] method no longer takes a filename. Instead the naming
  of the file has move to a [`workbook.save()`] method which replaces
  `workbook.close()`.
- There are now supporting [`workbook.save_to_path()`] and
  [`workbook.save_to_buffer()`] methods.

### Added

- Added new methods to get references to worksheet objects used by the workbook:

  - [`workbook.worksheet_from_name()`]
  - [`workbook.worksheet_from_index()`]
  - [`workbook.worksheets_mut()`]

- Made the [`Worksheet::new()`] method public and added the
  [`workbook.push_worksheet()`] to add Worksheet instances to a Workbook. See
  also the `rust_xlsxwriter` documentation on [Creating Worksheets] and working
  with the borrow checker.

[`Workbook::new()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.new
[`Worksheet::new()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.new

[`workbook.save()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.save
[Creating Worksheets]: https://rustxlsxwriter.github.io/worksheet/create.html
[`workbook.save_to_path()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.save_to_path
[`workbook.save_to_buffer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.save_to_buffer
[`workbook.worksheets_mut()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.worksheets_mut
[`workbook.push_worksheet()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.push_worksheet
[`workbook.worksheet_from_name()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.worksheet_from_name
[`workbook.worksheet_from_index()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.worksheet_from_index


## [0.8.0] - 2022-10-28

### Added

- Added support for creating files from paths via `workbook.new_from_path()`.

- Added support for creating file to a buffer via `workbook.new_from_buffer()` and `workbook.close_to_buffer()`.



## [0.7.0] - 2022-10-22

### Added

- Added an almost the complete set of Page Setup methods:

- Page Setup - Page

  - [`worksheet.set_portrait()`]
  - [`worksheet.set_landscape()`]
  - [`worksheet.set_print_scale()`]
  - [`worksheet.set_print_fit_to_pages()`]
  - [`worksheet.set_print_first_page_number()`]

[`worksheet.set_portrait()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_portrait
[`worksheet.set_landscape()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_landscape
[`worksheet.set_print_scale()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_scale
[`worksheet.set_print_fit_to_pages()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_fit_to_pages
[`worksheet.set_print_first_page_number()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_first_page_number

- Page Setup - Margins

  - [`worksheet.set_margins()`]
  - [`worksheet.set_print_center_horizontally()`]
  - [`worksheet.set_print_center_vertically()`]

[`worksheet.set_margins()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_margins
[`worksheet.set_print_center_horizontally()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_center_horizontally
[`worksheet.set_print_center_vertically()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_center_vertically


- Page Setup - Header/Footer

  - [`worksheet.set_header()`]
  - [`worksheet.set_footer()`]
  - [`worksheet.set_header_footer_scale_with_doc()`]
  - [`worksheet.set_header_footer_align_with_page()`]

[`worksheet.set_header()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_header
[`worksheet.set_footer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_footer
[`worksheet.set_header_footer_scale_with_doc()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_header_footer_scale_with_doc
[`worksheet.set_header_footer_align_with_page()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_header_footer_align_with_page

- Page Setup - Sheet

  - [`worksheet.set_print_area()`]
  - [`worksheet.set_repeat_rows()`]
  - [`worksheet.set_repeat_columns()`]
  - [`worksheet.set_print_gridlines()`]
  - [`worksheet.set_print_black_and_white()`]
  - [`worksheet.set_print_draft()`]
  - [`worksheet.set_print_headings()`]
  - [`worksheet.set_page_order()`]

[`worksheet.set_print_area()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_area
[`worksheet.set_repeat_rows()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_repeat_rows
[`worksheet.set_repeat_columns()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_repeat_columns
[`worksheet.set_print_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_gridlines
[`worksheet.set_print_black_and_white()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_black_and_white
[`worksheet.set_print_draft()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_draft
[`worksheet.set_print_headings()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_print_headings
[`worksheet.set_page_order()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_page_order

### Fixes

- Fix for cargo issue where chrono dependency had a RUSTSEC warning. [GitHub
  Issue #6].

[GitHub Issue #6]: https://github.com/jmcnamara/rust_xlsxwriter/issues/6

## [0.6.0] - 2022-10-18

### Added

- Added more page setup methods:

  - [`worksheet.set_header()`]
  - [`worksheet.set_footer()`]
  - [`worksheet.set_margins()`]

  See also the rust_xlsxwriter user documentation on [Adding Headers and
  Footers].

[`worksheet.set_header()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_header
[`worksheet.set_footer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_footer
[`worksheet.set_margins()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_margins
[Adding Headers and Footers]: https://rustxlsxwriter.github.io/worksheet/headers.html

## [0.5.0] - 2022-10-16

### Added

- Added page setup methods:

  - [`worksheet.set_zoom()`]
  - [`worksheet.set_landscape()`]
  - [`worksheet.set_paper_size()`]
  - [`worksheet.set_page_order()`]
  - [`worksheet.set_view_page_layout()`]
  - [`worksheet.set_view_page_break_preview()`]

[`worksheet.set_zoom()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_zoom
[`worksheet.set_paper_size()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_paper_size
[`worksheet.set_page_order()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_page_order
[`worksheet.set_landscape()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_landscape
[`worksheet.set_view_page_layout()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_view_page_layout
[`worksheet.set_view_page_break_preview()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.set_view_page_break_preview

## [0.4.0] - 2022-10-10

### Added

- Added support for array formulas and dynamic array formulas via
  [`worksheet.write_array()`] and
  [`worksheet.write_dynamic_array_formula_with_format()`].

See also the rust_xlsxwriter user documentation on [Dynamic Array support].

[`worksheet.write_array()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_array_formula
[`worksheet.write_dynamic_array_formula_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_dynamic_array_formula

[Dynamic Array support]: https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html

## [0.3.1] - 2022-10-01

### Fixed

- Fixed minor crate issue.


## [0.3.0] - 2022-10-01

### Added

- Added [`worksheet.write_boolean_with_format()`] method to support writing Excel boolean
  values.

[`worksheet.write_boolean_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_boolean

## [0.2.1] - 2022-09-22

### Fixed

- Fixed some minor crate/publishing issues.


## [0.2.0] - 2022-09-24

### Added

- First functional version. Supports the main data types and formatting.


## [0.1.0] - 2022-07-12

### Added

- Initial, non-functional crate, to initiate namespace.

