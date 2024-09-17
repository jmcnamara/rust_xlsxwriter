# Changelog

This is the changelog/release notes for the `rust_xlsxwriter` crate.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [0.77.0] - 2024-09-18

### Added

- Added support for [Chartsheets].

  A Chartsheet in Excel is a specialized type of worksheet that doesn't have
  cells but instead is used to display a single chart. It supports worksheet
  display options such as headers and footers, margins, tab selection and
  print properties.

  <img src="https://rustxlsxwriter.github.io/images/chartsheet.png">

  [Chartsheets]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/index.html#chartsheets


- Updated polars dependency to 0.43 to pick up latest Polars additions for
  [`polars_excel_writer`].


## [0.76.0] - 2024-09-11

### Added

  - Added support for adding Textbox shapes to worksheets. See the documentation
    for [`Shape`].

    <img src="https://rustxlsxwriter.github.io/images/app_textbox.png">


    [`Shape`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Shape.html


## [0.75.0] - 2024-09-02

### Removed

- Removed dependency on the `regex.rs` crate for smaller binary sizes. The only
  non-optional dependency is now `zip.rs`.

  An example of the size difference is shown below for one of the sample apps:

  | `app_hello_world` | v0.74.0 |v0.75.0 |
  |-------------------|---------|--------|
  | Debug             | 9.2M    | 4.2M   |
  | Release           | 3.4M    | 1.6M   |


- Removed the `Formula::use_future_functions()` and
  `Formula::use_table_functions()` methods since there functionality is now
  handled automatically as a result of the `regex` change.


## [0.74.0] - 2024-08-24

### Added

- Add methods to format cells separately from the data writing functions.

  In Excel the data in a worksheet cell is comprised of a type, a value and a
  format. When using `rust_xlsxwriter` the type is inferred and the value and
  format are generally written at the same time using methods like
  [`Worksheet::write_with_format()`].

  However, if required you can write the data separately and then add the format
  using the new methods like [`Worksheet::set_cell_format()`],
  [`Worksheet::set_range_format()`] and
  [`Worksheet::set_range_format_with_border()`].

  <img src="https://rustxlsxwriter.github.io/images/worksheet_set_range_format_with_border.png">

  [`Worksheet::set_cell_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_cell_format
  [`Worksheet::set_range_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_range_format
  [`Worksheet::write_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_with_format
  [`Worksheet::set_range_format_with_border()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_range_format_with_border

- Replaced the `IntoColor` trait with `Into<Color>` in all APIs. This doesn't
  require a change by the end user (unless they implemented `IntoColor` for
  their own type).

- Updated polars dependency to 0.42.0 to pick up latest Polars additions for
  [`polars_excel_writer`].



## [0.73.0] - 2024-08-02

### Added

- Added support for setting the default worksheet row height and also hiding all
  unused rows.

  <img src="https://rustxlsxwriter.github.io/images/worksheet_hide_unused_rows.png">

  See [`Worksheet::set_default_row_height()`] and  [`Worksheet::hide_unused_rows()`].

  [`Worksheet::hide_unused_rows()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.hide_unused_rows
  [`Worksheet::set_default_row_height()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_default_row_height


## [0.72.0] - 2024-07-26

### Added

  - Added support for cell Notes (previously called Comments). See the
    documentation for [`Note`].

    A Note is a post-it style message that is revealed when the user mouses over
    a worksheet cell. The presence of a Note is indicated by a small red
    triangle in the upper right-hand corner of the cell.

    <img src="https://rustxlsxwriter.github.io/images/app_notes.png">

    In versions of Excel prior to Office 365 Notes were referred to as
    "Comments". The name Comment is now used for a newer style threaded comment
    and Note is used for the older non threaded version.

    [`Note`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Note.html



## [0.71.0] - 2024-07-20

### Added

  - Added support for adding VBA Macros to `rust_xlsxwriter` using files
    extracted from Excel files.

    An Excel `xlsm` file is structurally the same as an `xlsx` file except that
    it contains an additional `vbaProject.bin` binary file containing VBA
    functions and/or macros.

    Unlike other components of an xlsx/xlsm file this data isn't stored in an
    XML format. Instead the functions and macros as stored as a pre-parsed
    binary format. As such it wouldn't be feasible to programmatically define
    macros and create a `vbaProject.bin` file from scratch (at least not in the
    remaining lifespan and interest levels of the author).

    Instead, as a workaround, the Rust [`vba_extract`] utility is used to
    extract `vbaProject.bin` files from existing xlsm files which can then be
    added to `rust_xlsxwriter` files.

    See [Working with VBA Macros].

    <img src="https://rustxlsxwriter.github.io/images/app_macros.png">

    [`vba_extract`]: https://crates.io/crates/vba_extract
    [Working with VBA Macros]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/macros/index.html


## [0.70.0] - 2024-07-14

### Added

  - Added support for adding Excel data validations to worksheet cells.

    Data validation is a feature of Excel that allows you to restrict the data
    that a user enters in a cell and to display associated help and warning
    messages. It also allows you to restrict input to values in a dropdown list.

    See [`DataValidation`] for details.

    <img src="https://rustxlsxwriter.github.io/images/data_validation_intro1.png">


    [`DataValidation`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.DataValidation.html


## [0.69.0] - 2024-07-01

### Added

  - Added support for adjusting the layout position of chart elements: plot
    area, legend, title and axis labels. See [`ChartLayout`].

   [`ChartLayout`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartLayout.html

### Fixed

  - Fixed issue where a worksheet name required quoting when used with
    `Worksheet::repeat_row()`. There was some checks to handle this but they
    weren't comprehensive enough. [Issue #95].

    [Issue #95]: https://github.com/jmcnamara/rust_xlsxwriter/issues/95


## [0.68.0] - 2024-06-18

### Added

  - Added support for urls in images. [Feature Request #91].

    [Feature Request #91]: https://github.com/jmcnamara/rust_xlsxwriter/issues/91

### Changed

  - Changed the method signatures of the [`Image`] helper methods from `&mut
    self` to `mut self` to allow method chaining. This is an API/ABI break.


## [0.67.0] - 2024-06-17

### Added

  - Updated the default `zip.rs` requirement to v2+ to pick up a fix for
    [zip-rs/zip2#100] when dealing with 64k+ internal files in an xlsx
    container. As a result of this, `rust_xlsxwriter` now has a matching `msrv`
    (Minimum Supported Rust Version) of v1.73.0.

  - Replaced the dependency on `lazy_static` with `std::cell::OnceLock`. The
    only non-optional requirements are now `zip` and `regex`. This was made
    possible by the above `msrv` update. See [Feature Request #24].

  - Added an optional dependency on the [ryu] crate to speed up writing large
    amounts of worksheet numeric data. The feature flag is `ryu`.

    This feature has a benefit when writing more than 300,000 numeric data
    cells. When writing 5,000,000 numeric cells it can be 30% faster than the
    standard feature set. See the following [performance analysis] but also
    please test it for your own scenario when enabling it since a performance
    improvement is not guaranteed in all cases.

  - Added Excel [Sensitivity Label] cookbook example and explanation.

    Sensitivity Labels are a property that can be added to an Office 365
    document to indicate that it is compliant with a company’s information
    protection policies. Sensitivity Labels have designations like
    “Confidential”, “Internal use only”, or “Public” depending on the policies
    implemented by the company. They are generally only enabled for enterprise
    versions of Office.

  - Updated all dependency versions to the latest current values.

    [ryu]: https://docs.rs/ryu/latest/ryu/index.html
    [zip-rs/zip2#100]: https://github.com/zip-rs/zip2/issues/100
    [Sensitivity Label]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html#document-properties-setting-the-sensitivity-label
    [Feature Request #24]: https://github.com/jmcnamara/rust_xlsxwriter/issues/24
    [performance analysis]: https://github.com/jmcnamara/rust_xlsxwriter/issues/93


## [0.66.0] - 2024-06-12

### Added

  - Added example of using a secondary X axis. See [Chart Secondary Axes].

### Changed

  - Changed `ChartSeries::set_y2_axis()` to `ChartSeries::set_secondary_axis()` for API consistency.

   [`ChartSeries::set_secondary_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_secondary_axis


## [0.65.0] - 2024-06-11

### Added

  - Added support for [Chart Secondary Axes] and [Combined Charts].

    <img src="https://rustxlsxwriter.github.io/images/chart_series_set_secondary_axis.png">

    [Chart Secondary Axes]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/index.html#secondary-axes
    [Combined Charts]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/index.html#combined-charts


## [0.64.2] - 2024-04-13

### Fixed

  - Fixed internal links in table of contents.


## [0.64.1] - 2024-03-26

### Added

  - Added the [`worksheet::set_screen_gridlines()`] method to turn on/offscreen gridlines.

  [`worksheet::set_screen_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_screen_gridlines

  - Added updated docs on [Working with Workbooks] and [Working with Worksheets].

  [Working with Workbooks]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/index.html
  [Working with Worksheets]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/index.html


## [0.64.0] - 2024-03-18

### Added

- Add support for worksheet sparklines. Sparklines are a feature of Excel 2010+
  which allows you to add small charts to worksheet cells. These are useful for
  showing data trends in a compact visual format.

  See [Working with Sparklines].

  <img src="https://rustxlsxwriter.github.io/images/sparklines1.png">

  [Working with Sparklines]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/sparkline/index.html


## [0.63.0] - 2024-02-25

### Added

- Added support for embedding images into worksheets with
  [`worksheet::embed_image()`] and [`worksheet::embed_image_with_format()`] and
  the [`Image`] struct. See the [Embedded Images] example.

  This can be useful if you are building up a spreadsheet of products with a
  column of images for each product. Embedded images move with the cell so they
  can be used in worksheet tables or data ranges that will be sorted or
  filtered.

  This functionality is the equivalent of Excel's menu option to insert an image
  using the option to "Place in Cell" which is available in Excel 365 versions
  from 2023 onwards.

  [`worksheet::embed_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.embed_image
  [`worksheet::embed_image_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.embed_image_with_format

  [Embedded Images]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html#insert-images-embedding-an-image-in-a-cell


- Updated polars dependency to 0.37.2 to pick up latest Polars additions for
  [`polars_excel_writer`].


- Added [`utility::check_sheet_name()`] function to allow checking for valid
  worksheet names according to Excel's naming rules. This functionality was
  previously `pub(crate)` private.

  [Feature Request #83].

  [Feature Request #83]: https://github.com/jmcnamara/rust_xlsxwriter/issues/83

  [`utility::check_sheet_name()`]:
      https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.check_sheet_name.html

### Removed

- Removed unnecessary lifetime on [`Format`] objects used in Worksheet `write()`
  methods. This allows the the [`IntoExcelData`] trait to be defined for user
  types and have them include a default format. See [Feature Request #85].

  [Feature Request #85]: https://github.com/jmcnamara/rust_xlsxwriter/issues/85
  [`Format`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html

## [0.62.0] - 2024-01-24

### Added

- Added support for adding a worksheet [`Table`] as a serialization format. See [`SerializeFieldOptions::set_table()`].

- Added [`Worksheet::get_serialize_dimensions()`] and
  [`Worksheet::get_serialize_column_dimensions()`] methods to get dimensions
  from a serialized range.

- Updated polars dependency to 0.36.2 to pick up Polars `AnyData` changes for
  [`polars_excel_writer`].


[`Worksheet::get_serialize_dimensions()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.get_serialize_dimensions

[`Worksheet::get_serialize_column_dimensions()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.get_serialize_column_dimensions


[`SerializeFieldOptions::set_table()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/struct.SerializeFieldOptions.html#method.set_table


### Changed

- Changed APIs for [`Table`] to return `Table` instead of `&Table` to allow
  methods to be chained. This makes worksheet Table usage easier during
  serialization. Note that this is a backward incompatible change.



## [0.61.0] - 2024-01-13

### Added

- Added support for a `XlsxSerialize` derive and struct attributes to control
  the formatting and options of the Excel output during serialization. These are
  similar in intention to Serde container/field attributes.

  See [Controlling Excel output via `XlsxSerialize` and struct attributes] and
  [Working with Serde].

  [Feature Request #66].

  [Feature Request #66]: https://github.com/jmcnamara/rust_xlsxwriter/issues/66

  [Controlling Excel output via `XlsxSerialize` and struct attributes]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html#controlling-excel-output-via-xlsxserialize-and-struct-attributes


- Added `XlsxError::CustomError` as a target error for external crates/apps.

  [Feature Request #72].

  [Feature Request #72]: https://github.com/jmcnamara/rust_xlsxwriter/issues/72


## [0.60.0] - 2024-01-02

### Added

- Added support for setting Serde headers using deserialization of a target
  struct type as well as the previous method of using serialization and an
  instance of the struct type. See [Setting serialization headers].

  [Feature Request #63].

- Added additional support for serialization header and field options via
  [`CustomSerializeField`].

- Added support for writing `Result<T, E>` with [`Worksheet::write()`] when `T`
  and `E` are supported types.

  [Feature Request #64].

[Feature Request #63]: https://github.com/jmcnamara/rust_xlsxwriter/issues/63
[Feature Request #64]: https://github.com/jmcnamara/rust_xlsxwriter/issues/64
[Setting serialization headers]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/index.html#setting-serialization-headers
[`CustomSerializeField`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/serializer/struct.CustomSerializeField.html

## [0.59.0] - 2023-12-15

### Added

- Added [`serialize_chrono_option_naive_to_excel()`] to help serialization of
  `Option` Chrono types. [Feature Request #62].

[Feature Request #62]: https://github.com/jmcnamara/rust_xlsxwriter/issues/62

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

- Added support for `Option<T>` wrapped types to [`Worksheet::write()`].

  [Feature Request #59].

[Feature Request #59]: https://github.com/jmcnamara/rust_xlsxwriter/issues/59


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

[`ChartDataTable`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartDataTable.html
[`Chart::set_data_table()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.set_data_table
[`ChartAxis::set_crossing()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_crossing
[`ChartAxis::set_label_alignment()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_label_alignment
[`ChartAxis::set_display_unit_type()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_display_unit_type
[`ChartSeries::set_smooth()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_smooth


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

[`Chart::set_drop_lines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.set_drop_lines
[`Chart::set_up_down_bars()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.set_up_down_bars
[`Chart::set_high_low_lines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.set_high_low_lines

[`ChartAxis::set_date_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_date_axis
[`ChartAxis::set_text_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_text_axis
[`ChartAxis::set_automatic_axis()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_automatic_axis

[`ChartAxis::set_min_date()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_min_date
[`ChartAxis::set_max_date()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_max_date

[`Worksheet::write_row_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_row_with_format
[`Worksheet::write_column_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_column_with_format


[Stock Chart]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html#chart-stock-excel-stock-chart-example

## [0.52.0] - 2023-10-20

### Added

- Added support for chart series error bars via the [`ChartErrorBars`] struct
  and methods.

[`ChartErrorBars`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartErrorBars.html

### Fixed

- Fixed XML error in non-Pie charts.

  [GitHub Issue #55].

[GitHub Issue #55]: https://github.com/jmcnamara/rust_xlsxwriter/issues/55



## [0.51.0] - 2023-10-15

### Added

- Added support for chart gradient fill formatting via the [`ChartGradientFill`] struct and methods.

- Added support for formatting the chart trendlines data labels via the
  [`ChartTrendline::set_label_font`] and [`ChartTrendline::set_label_format`].


[`ChartGradientFill`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartGradientFill.html
[`ChartTrendline::set_label_font`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartTrendline.html#method.set_label_font
[`ChartTrendline::set_label_format`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartTrendline.html#method.set_label_format


## [0.50.0] - 2023-10-12

### Added

- Added support for chart trendlines (Linear, Polynomial, Moving Average, etc.)
  via the [`ChartTrendline`] struct and methods.

- Added the [`Worksheet::set_very_hidden()`] method to hide a worksheet similar
  to the [`Worksheet::set_hidden()`] method. The difference is that the worksheet
  can only be unhidden by VBA and cannot be unhidden in the the Excel user
  interface.

- Added support for leader lines to non-Pie charts.

### Fixed

- Fixed handling of [future functions] in table formulas.


[`ChartTrendline`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartTrendline.html
[`Worksheet::set_very_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_very_hidden
[future functions]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Formula.html#formulas-added-in-excel-2010-and-later


## [0.49.0] - 2023-09-19

### Added

- Added chart options to control how non-data cells are displayed.

  - [`Chart::show_empty_cells_as()`]
  - [`Chart::show_na_as_empty_cell()`]
  - [`Chart::show_hidden_data()`]

[`Chart::show_hidden_data()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.show_hidden_data
[`Chart::show_empty_cells_as()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.show_empty_cells_as
[`Chart::show_na_as_empty_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.show_na_as_empty_cell


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

- Added [`Workbook::save_to_writer()`] method to make it easier to interact with
  interfaces that implement the `<W: Write>` trait.

[rust_xlsx_wasm_example]: https://github.com/Clipi-12/rust_xlsx_wasm_example
[`Workbook::save_to_writer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.save_to_writer


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
  [`TableColumn::set_format()`] and [`TableColumn::set_header_format()`] methods.

[`TableColumn::set_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.TableColumn.html#method.set_format
[`TableColumn::set_header_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.TableColumn.html#method.set_header_format


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

[`cell_range()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.cell_range.html
[`row_col_to_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.row_col_to_cell.html
[`cell_range_absolute()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.cell_range_absolute.html
[`column_number_to_name()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.column_number_to_name.html
[`column_name_to_number()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.column_name_to_number.html
[`row_col_to_cell_absolute()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/utility/fn.row_col_to_cell_absolute.html

[GitHub Issue #45]: https://github.com/jmcnamara/rust_xlsxwriter/issues/45
[GitHub Issue #47]: https://github.com/jmcnamara/rust_xlsxwriter/issues/47

[Tutorial]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/tutorial/index.html
[Cookbook]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/cookbook/index.html
[`write_datetime()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_datetime
[`write_datetime_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_datetime_with_format


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
  [`Worksheet::write_datetime()`]. This is a soft deprecation.

[`Chrono`]: https://docs.rs/chrono/latest/chrono
[`ExcelDateTime`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html
[`IntoExcelDateTime`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/trait.IntoExcelDateTime.html
[`Worksheet::write_datetime()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_datetime


## [0.40.0] - 2023-05-31

- Added support for worksheet tables. See [`Table`] and the
  [`Worksheet::add_table()`] method.

- Added support for the `bool` type in the generic [`Worksheet::write()`] method.

[`Table`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Table.html
[`Worksheet::add_table()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.add_table


## [0.39.0] - 2023-05-23

### Added

- Added [`Worksheet::write_row()`], [`Worksheet::write_column()`],
  [`Worksheet::write_row_matrix()`] and [`Worksheet::write_column_matrix()`]
  methods to write arrays/iterators of data.

- Added [`Formula`] and [`Url`] types to use with generic [`Worksheet::write()`].

  [Feature Request #16].

- Make several string handling APIs more generic using `impl Into<String>`.

  [Feature Request #16].

- Renamed/refactored `XlsxColor` to [`Color`] for API consistency. The
  `XlsxColor` type alias is still available for backward compatibility.

[`Url`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Url.html
[`Color`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.Color.html
[`Formula`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Formula.html
[Feature Request #16]: https://github.com/jmcnamara/rust_xlsxwriter/issues/16
[`Worksheet::write()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write
[`Worksheet::write_row()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_row
[`Worksheet::write_column()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_column
[`Worksheet::write_row_matrix()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_row_matrix
[`Worksheet::write_column_matrix()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_column_matrix


## [0.38.0] - 2023-05-05

### Added

- Added several Chart axis and series methods:

  - [`ChartAxis::set_hidden()`]
  - [`ChartAxis::set_label_interval()`]
  - [`ChartAxis::set_label_position()`]
  - [`ChartAxis::set_log_base()`]
  - [`ChartAxis::set_major_gridlines()`]
  - [`ChartAxis::set_major_gridlines_line()`]
  - [`ChartAxis::set_major_tick_type()`]
  - [`ChartAxis::set_major_unit()`]
  - [`ChartAxis::set_max()`]
  - [`ChartAxis::set_min()`]
  - [`ChartAxis::set_minor_gridlines()`]
  - [`ChartAxis::set_minor_gridlines_line()`]
  - [`ChartAxis::set_minor_tick_type()`]
  - [`ChartAxis::set_minor_unit()`]
  - [`ChartAxis::set_position_between_ticks()`]
  - [`ChartAxis::set_reverse()`]
  - [`ChartAxis::set_tick_interval()`]
  - [`ChartSeries::set_invert_if_negative()`]
  - [`ChartSeries::set_invert_if_negative_color()`]

[`ChartAxis::set_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_hidden
[`ChartAxis::set_label_interval()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_label_interval
[`ChartAxis::set_label_position()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_label_position
[`ChartAxis::set_log_base()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_log_base
[`ChartAxis::set_major_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_major_gridlines
[`ChartAxis::set_major_gridlines_line()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_major_gridlines_line
[`ChartAxis::set_major_tick_type()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_major_tick_type
[`ChartAxis::set_major_unit()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_major_unit
[`ChartAxis::set_max()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_max
[`ChartAxis::set_min()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_min
[`ChartAxis::set_minor_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_minor_gridlines
[`ChartAxis::set_minor_gridlines_line()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_minor_gridlines_line
[`ChartAxis::set_minor_tick_type()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_minor_tick_type
[`ChartAxis::set_minor_unit()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_minor_unit
[`ChartAxis::set_position_between_ticks()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_position_between_ticks
[`ChartAxis::set_reverse()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_reverse
[`ChartAxis::set_tick_interval()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartAxis.html#method.set_tick_interval
[`ChartSeries::set_invert_if_negative()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_invert_if_negative
[`ChartSeries::set_invert_if_negative_color()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_invert_if_negative_color


## [0.37.0] - 2023-04-30

### Added

- Added font formatting support to chart titles, legends, axes and data labels
  via [`ChartFont`] and various `set_font()` methods.

- Made [`Worksheet::write_string()`] and [`Worksheet::write()`] more generic via
  `impl Into<String>` to allow them to handle `&str`, `&String`, `String`, and
  `Cow<>` types.

  See [GitHub Feature Request #35].

[`ChartFont`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartFont.html
[`Worksheet::write()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_string
[`Worksheet::write_string()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_string
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
  [`ChartDataLabel`], [`Chart::series.set_data_label()`] and [`Chart::series.set_custom_data_labels()`].

[`ChartDataLabel`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartDataLabel.html
[`Chart::series.set_data_label()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_data_label
[`Chart::series.set_custom_data_labels()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_custom_data_labels


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

- Added [`Chart::set_rotation()`] and [`Chart::set_hole_size()`] methods for Pie and Doughnut charts.

- Added support to differentiate between `Color::Default` and
  `Color::Automatic` colors for Excel elements. These are usually equivalent
  but there are some cases where the "Automatic" color, which can be set at a
  system level, is different from the Default color.

[Pie Chart]: https://rustxlsxwriter.github.io/examples/pie_chart.html
[`ChartPoint`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartPoint.html
[`ChartMarker`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartMarker.html
[`Chart::set_rotation()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.set_rotation
[`Chart::set_hole_size()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html#method.set_hole_size


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

- Added `IntoColor` trait to allow syntactic shortcuts for [`Color`]
  parameters in methods. So now you can set a RGB color like this
  `object.set_color("#FF7F50")` instead of the more verbose
  `object.set_color(Color::RGB(0xFF7F50))`. This addition doesn't require
  any API changes from the end user.

- Added [`Worksheet::insert_image_fit_to_cell()`] method to add an image to a
  worksheet and scale it so that it fits in a cell. This method can be useful
  when creating a product spreadsheet with a column of images for each product.

  See also the [insert_image_to_fit] example in the user guide.

- Added [`Chart::series.set_gap()`] and [`Chart::series.set_overlap()`] method to control layout
  of histogram style charts.

[`ChartFormat`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartFormat.html
[Chart Fill Pattern]: https://rustxlsxwriter.github.io/examples/chart_pattern.html
[insert_image_to_fit]: https://rustxlsxwriter.github.io/examples/insert_image_to_fit.html
[`Chart::series.set_gap()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_gap
[`Chart::series.set_overlap()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartSeries.html#method.set_overlap
[`Worksheet::insert_image_fit_to_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.insert_image_fit_to_cell


## [0.29.0] - 2023-03-16

### Added

- Added support for resizing and object positioning to the [`Chart`] struct.

- Added handling for `chrono` date/time types to the generic
  [`Worksheet::write()`] method.


## [0.28.0] - 2023-03-14

### Added

- Added support for positioning or hiding Chart legends. See [`ChartLegend`].

[`ChartLegend`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.ChartLegend.html


## [0.27.0] - 2023-03-13

### Added

- Added support for Charts via the [`Chart`] struct and the
  [`Worksheet::insert_chart()`] method. See also the [Chart Examples] in the user
  guide.

- Added a generic [`Worksheet::write()`] method that writes string or number
  types. This will be extended in an upcoming release to provide a single
  `write()` method for all of the currently supported types.

  It also allows the user to extend [`Worksheet::write()`] to handle user defined
  types via the [`IntoExcelData`] trait. See also the [Writing Generic data]
  example in the user guide.

[`Chart`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/chart/struct.Chart.html
[Chart Examples]: https://rustxlsxwriter.github.io/examples/simple_chart.html
[`IntoExcelData`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/trait.IntoExcelData.html
[`Worksheet::write()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write
[Writing Generic data]: https://rustxlsxwriter.github.io/examples/generic_write.html
[`Worksheet::insert_chart()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.insert_chart


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

- Added ability to filter columns in [`Worksheet::autofilter()`] ranges via
  [`Worksheet::filter_column()`] and [`FilterCondition`].

  The library automatically hides any rows that don't match the supplied
  criteria. This is an additional feature that isn't available in the other
  language ports of "xlsxwriter".

  See also the [Working with Autofilters] section of the Users Guide.

[`FilterCondition`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.FilterCondition.html
[Working with Autofilters]: https://rustxlsxwriter.github.io/formulas/autofilters.html
[`Worksheet::autofilter()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.autofilter
[`Worksheet::filter_column()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.filter_column


## [0.24.0] - 2023-01-18

### Added

- Added support for hiding rows and columns (to hide intermediate calculations)
  via the [`Worksheet::set_column_hidden()`] and[`Worksheet::set_row_hidden()`]
  method. This is also a required precursor to adding autofilter conditions.
- Added the [ObjectMovement] enum to control how a worksheet object, such a
  an image, moves when the cells underneath it are moved, resized or deleted.

[ObjectMovement]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.ObjectMovement.html
[`Worksheet::set_row_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_row_hidden
[`Worksheet::set_column_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_column_hidden

## [0.23.0] - 2023-01-16

### Added

Added more page setup methods.

- Added [`Worksheet::set_selection()`] method to select a cell or range of cells in a worksheet.
- Added [`Worksheet::set_top_left_cell()`] method to set the top and leftmost visible cell.
- Added [`Worksheet::set_page_breaks()`] method to add page breaks to a worksheet.

[`Worksheet::set_selection()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_selection
[`Worksheet::set_page_breaks()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_page_breaks
[`Worksheet::set_top_left_cell()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_top_left_cell


## [0.22.0] - 2023-01-13

### Added

- Added support for worksheet protection via the [`Worksheet::protect()`],
  [`Worksheet::protect_with_password()`] and [`Worksheet::protect_with_options()`].

  See also the section on [Worksheet protection] in the user guide.

- Add option to make the xlsx file read-only when opened by Excel via the
  [`Workbook::read_only_recommended()`] method.


[Worksheet protection]:  https://rustxlsxwriter.github.io/worksheet/protection.html
[`Worksheet::protect()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.protect
[`Worksheet::protect_with_options()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.protect_with_options
[`Workbook::read_only_recommended()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.read_only_recommended
[`Worksheet::protect_with_password()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.protect_with_password


## [0.21.0] - 2023-01-09

### Added

- Added support for setting document metadata properties such as Author and
  Creation Date. For more details see [`DocProperties`] and
  [`workbook::set_properties()`].

[`DocProperties`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.DocProperties.html
[`workbook::set_properties()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.set_properties

### Changed

- Change date/time parameters to references in [`Worksheet::write_datetime()`],
  `worksheet.write_date()` and `worksheet.write_time()` for consistency.

[`Worksheet::write_datetime()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_datetime


## [0.20.0] - 2023-01-06

### Added

- Improved fitting algorithm for [`Worksheet::autofit()`]. See also the
  [app_autofit] sample application.

### Changed

- The `worksheet.set_autofit()` method has been renamed to `worksheet.autofit()`
  for consistency with the other language versions of this library.


## [0.19.0] - 2022-12-27

### Added

- Added support for created defined variable names at a workbook and worksheet
  level via [`Workbook::define_name()`].

  See also [Using defined names] in the user guide.

- Added initial support for autofilters via [`Worksheet::autofilter()`].

  Note, adding filter criteria isn't currently supported. That will be added in
  an upcoming version. See also [Adding Autofilters] in the user guide.

[Adding Autofilters]: https://rustxlsxwriter.github.io/examples/autofilter.html
[Using defined names]: https://rustxlsxwriter.github.io/examples/defined_names.html
[`Worksheet::autofilter()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.autofilter
[`Workbook::define_name()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.define_name


## [0.18.0] - 2022-12-19

### Added

- Added support for "rich" strings with multiple font formats via
  [`Worksheet::write_rich_string()`] and [`Worksheet::write_rich_string_with_format()`].
  For example strings like "This is **bold** and this is *italic*".

  See also the [Rich strings example] in the user guide.

[Rich strings example]: https://rustxlsxwriter.github.io/examples/rich_strings.html
[`Worksheet::write_rich_string_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_rich_string_with_format
[`Worksheet::write_rich_string()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_rich_string

## [0.17.1] - 2022-12-18

### Fixed

- Fixes issue where header image files became corrupt during incremental saves.
  Also fixes similar issues in some formatting code.


## [0.17.0] - 2022-12-17

### Added

- Added support for images in headers/footers via the
  [`Worksheet::set_header_image()`] and [`Worksheet::set_footer_image()`] methods.

  See the [Headers and Footers] and [Adding a watermark] examples in the user guide.

[Headers and Footers]: https://rustxlsxwriter.github.io/examples/headers.html
[Adding a watermark]: https://rustxlsxwriter.github.io/examples/watermark.html
[`Worksheet::set_footer_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_footer_image
[`Worksheet::set_header_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_header_image


## [0.16.0] - 2022-12-09

### Added

- Replicate the optimization used by Excel where it only stores one copy of a
  repeated/duplicate image in a workbook.


## [0.15.0] - 2022-12-08

### Added

- Added support for images in buffers via [`Image::new_from_buffer()`].

- Added image accessability features via [`Image::set_alt_text()`] and[`Image::set_decorative()`].

[`Image::set_alt_text()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Image.html#method.set_alt_text
[`Image::set_decorative()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Image.html#method.set_decorative
[`Image::new_from_buffer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Image.html#method.new_from_buffer


## [0.14.0] - 2022-12-05

### Added

- Added support for inserting images into worksheets with
  [`Worksheet::insert_image()`] and [`Worksheet::insert_image_with_offset()`] and
  the [`Image`] struct.

  See also the [images example] in the user guide.

  Upcoming versions of the library will support additional image handling
  features such as EMF and WMF formats, removal of duplicate images, hyperlinks
  in images and images in headers/footers.

### Removed

- The [`Workbook::save()`] method has been extended to handle paths or strings.
  The `workbook.save_to_path()` method has been removed. See [PR #15].

[`Image`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html
[PR #15]: https://github.com/jmcnamara/rust_xlsxwriter/pull/15
[images example]: https://rustxlsxwriter.github.io/examples/images.html
[`Worksheet::insert_image()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.insert_image
[`Worksheet::insert_image_with_offset()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.insert_image_with_offset


## [0.13.0] - 2022-11-21

### Added

- Added support for writing hyperlinks in worksheets via the following methods:

  - [`Worksheet::write_url()`] to write a link with a default hyperlink format style.
  - [`Worksheet::write_url_with_text()`] to add alternative text to the link.
  - [`Worksheet::write_url_with_format()`] to add an alternative format to the link.

See also the [hyperlinks example] in the user guide.

[hyperlinks example]: https://rustxlsxwriter.github.io/examples/hyperlinks.html
[`Worksheet::write_url()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_url
[`Worksheet::write_url_with_text()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_url_with_text
[`Worksheet::write_url_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_url_with_format


## [0.12.1] - 2022-11-09

### Changed

- Dependency changes to make WASM compilation easier:

  - Reduced the `zip` dependency to the minimum import only.
  - Removed dependency on `tempfile`. The library now uses in memory files.

## [0.12.0] - 2022-11-06

### Added

- Added [`Worksheet::merge_range()`] method.
- Added support for Theme colors to [`Color`]. See also [Working with
  Colors] in the user guide.

[`Color`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/enum.Color.html
[Working with Colors]: https://rustxlsxwriter.github.io/colors/intro.html
[`Worksheet::merge_range()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.merge_range


## [0.11.0] - 2022-11-04

### Added

- Added several worksheet methods for working with worksheet tabs:

  - [`Worksheet::set_active()`]: Set the active/visible worksheet.
  - [`Worksheet::set_tab_color()`]: Set the tab color.
  - [`Worksheet::set_hidden()`]: Hide a worksheet.
  - [`Worksheet::set_selected()`]: Set a worksheet as selected.
  - [`Worksheet::set_first_tab()`]: Set the first visible tab.

  See also [Working with worksheet tabs] in the user guide.

[`Worksheet::set_active()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_active
[`Worksheet::set_hidden()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_hidden
[`Worksheet::set_selected()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_selected
[`Worksheet::set_tab_color()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_tab_color
[`Worksheet::set_first_tab()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_first_tab
[Working with worksheet tabs]: https://rustxlsxwriter.github.io/worksheet/tabs.html


## [0.10.0] - 2022-11-03

### Added

- Added a simulated [`Worksheet::autofit()`] method to automatically adjust
  the width of columns with data. See also the [app_autofit] sample application.

- Added the [`Worksheet::set_freeze_panes()`] method to set "freeze" panes for
  worksheets. See also the [app_panes] example application.

[app_panes]: https://rustxlsxwriter.github.io/examples/panes.html
[app_autofit]: https://rustxlsxwriter.github.io/examples/autofit.html
[`Worksheet::autofit()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.autofit
[`Worksheet::set_freeze_panes()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_freeze_panes

## [0.9.0] - 2022-10-31

**Note, this version contains a major backward incompatible API change where it
restructures the Workbook constructor/destructor sequence and introduces a
`save()` method to replace `close()`.**

### Changed

- The [`Workbook::new()`] method no longer takes a filename. Instead the naming
  of the file has move to a [`Workbook::save()`] method which replaces
  `workbook.close()`.


### Added

- Added new methods to get references to worksheet objects used by the workbook:

  - [`Workbook::worksheet_from_name()`]
  - [`Workbook::worksheet_from_index()`]
  - [`Workbook::worksheets_mut()`]

- Made the [`Worksheet::new()`] method public and added the
  [`Workbook::push_worksheet()`] to add Worksheet instances to a Workbook. See
  also the `rust_xlsxwriter` documentation on [Creating Worksheets] and working
  with the borrow checker.

[`Workbook::new()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.new
[`Worksheet::new()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.new

[`Workbook::save()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.save
[Creating Worksheets]: https://rustxlsxwriter.github.io/worksheet/create.html
[`Workbook::worksheets_mut()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.worksheets_mut
[`Workbook::push_worksheet()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.push_worksheet
[`Workbook::worksheet_from_name()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.worksheet_from_name
[`Workbook::worksheet_from_index()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/workbook/struct.Workbook.html#method.worksheet_from_index


## [0.8.0] - 2022-10-28

### Added

- Added support for creating files from paths via `workbook.new_from_path()`.

- Added support for creating file to a buffer via `workbook.new_from_buffer()` and `workbook.close_to_buffer()`.



## [0.7.0] - 2022-10-22

### Added

- Added an almost the complete set of Page Setup methods:

- Page Setup - Page

  - [`Worksheet::set_portrait()`]
  - [`Worksheet::set_landscape()`]
  - [`Worksheet::set_print_scale()`]
  - [`Worksheet::set_print_fit_to_pages()`]
  - [`Worksheet::set_print_first_page_number()`]

[`Worksheet::set_portrait()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_portrait
[`Worksheet::set_landscape()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_landscape
[`Worksheet::set_print_scale()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_scale
[`Worksheet::set_print_fit_to_pages()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_fit_to_pages
[`Worksheet::set_print_first_page_number()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_first_page_number

- Page Setup - Margins

  - [`Worksheet::set_margins()`]
  - [`Worksheet::set_print_center_horizontally()`]
  - [`Worksheet::set_print_center_vertically()`]

[`Worksheet::set_margins()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_margins
[`Worksheet::set_print_center_horizontally()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_center_horizontally
[`Worksheet::set_print_center_vertically()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_center_vertically


- Page Setup - Header/Footer

  - [`Worksheet::set_header()`]
  - [`Worksheet::set_footer()`]
  - [`Worksheet::set_header_footer_scale_with_doc()`]
  - [`Worksheet::set_header_footer_align_with_page()`]

[`Worksheet::set_header()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_header
[`Worksheet::set_footer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_footer
[`Worksheet::set_header_footer_scale_with_doc()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_header_footer_scale_with_doc
[`Worksheet::set_header_footer_align_with_page()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_header_footer_align_with_page

- Page Setup - Sheet

  - [`Worksheet::set_print_area()`]
  - [`Worksheet::set_repeat_rows()`]
  - [`Worksheet::set_repeat_columns()`]
  - [`Worksheet::set_print_gridlines()`]
  - [`Worksheet::set_print_black_and_white()`]
  - [`Worksheet::set_print_draft()`]
  - [`Worksheet::set_print_headings()`]
  - [`Worksheet::set_page_order()`]

[`Worksheet::set_print_area()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_area
[`Worksheet::set_repeat_rows()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_repeat_rows
[`Worksheet::set_repeat_columns()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_repeat_columns
[`Worksheet::set_print_gridlines()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_gridlines
[`Worksheet::set_print_black_and_white()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_black_and_white
[`Worksheet::set_print_draft()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_draft
[`Worksheet::set_print_headings()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_print_headings
[`Worksheet::set_page_order()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_page_order

### Fixes

- Fix for cargo issue where chrono dependency had a RUSTSEC warning. [GitHub
  Issue #6].

[GitHub Issue #6]: https://github.com/jmcnamara/rust_xlsxwriter/issues/6

## [0.6.0] - 2022-10-18

### Added

- Added more page setup methods:

  - [`Worksheet::set_header()`]
  - [`Worksheet::set_footer()`]
  - [`Worksheet::set_margins()`]

  See also the `rust_xlsxwriter` user documentation on [Adding Headers and
  Footers].

[`Worksheet::set_header()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_header
[`Worksheet::set_footer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_footer
[`Worksheet::set_margins()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_margins
[Adding Headers and Footers]: https://rustxlsxwriter.github.io/worksheet/headers.html

## [0.5.0] - 2022-10-16

### Added

- Added page setup methods:

  - [`Worksheet::set_zoom()`]
  - [`Worksheet::set_landscape()`]
  - [`Worksheet::set_paper_size()`]
  - [`Worksheet::set_page_order()`]
  - [`Worksheet::set_view_page_layout()`]
  - [`Worksheet::set_view_page_break_preview()`]

[`Worksheet::set_zoom()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_zoom
[`Worksheet::set_paper_size()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_paper_size
[`Worksheet::set_page_order()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_page_order
[`Worksheet::set_landscape()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_landscape
[`Worksheet::set_view_page_layout()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_view_page_layout
[`Worksheet::set_view_page_break_preview()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_view_page_break_preview

## [0.4.0] - 2022-10-10

### Added

- Added support for array formulas and dynamic array formulas via
  [`Worksheet::write_array()`] and
  [`Worksheet::write_dynamic_array_formula_with_format()`].

See also the `rust_xlsxwriter` user documentation on [Dynamic Array support].

[`Worksheet::write_array()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_array_formula
[`Worksheet::write_dynamic_array_formula_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_dynamic_array_formula

[Dynamic Array support]: https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html

## [0.3.1] - 2022-10-01

### Fixed

- Fixed minor crate issue.


## [0.3.0] - 2022-10-01

### Added

- Added [`Worksheet::write_boolean_with_format()`] method to support writing Excel boolean
  values.

[`Worksheet::write_boolean_with_format()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_boolean

## [0.2.1] - 2022-09-22

### Fixed

- Fixed some minor crate/publishing issues.


## [0.2.0] - 2022-09-24

### Added

- First functional version. Supports the main data types and formatting.


## [0.1.0] - 2022-07-12

### Added

- Initial, non-functional crate, to initiate namespace.

