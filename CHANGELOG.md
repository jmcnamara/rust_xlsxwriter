# Changelog

All notable changes to rust_xlsxwriter will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.8.0] - 2022-10-28

### Added

- Added support for creating files from paths via [`workbook.new_from_path()`].

- Added support for creating file to a buffer via [`workbook.new_from_buffer()`] and [`workbook.close_to_buffer()`].


[`workbook.new_from_path()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.new_from_path
[`workbook.new_from_buffer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.new_from_buffer
[`workbook.close_to_buffer()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Workbook.html#method.close_to_buffer


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
  [`worksheet.write_array_formula()`] and
  [`worksheet.write_dynamic_array_formula()`].

See also the rust_xlsxwriter user documentation on [Dynamic Array support].

[`worksheet.write_array_formula()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_array_formula
[`worksheet.write_dynamic_array_formula()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_dynamic_array_formula

[Dynamic Array support]: https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html

## [0.3.1] - 2022-10-01

### Fixed

- Fixed minor crate issue.


## [0.3.0] - 2022-10-01

### Added

- Added [`worksheet.write_boolean()`] method to support writing Excel boolean
  values.

[`worksheet.write_boolean()`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.write_boolean

## [0.2.1] - 2022-09-22

### Fixed

- Fixed some minor crate/publishing issues.


## [0.2.0] - 2022-09-24

### Added

- First functional version. Supports the main data types and formatting.


## [0.1.0] - 2022-07-12

### Added

- Initial, non-functional crate, to initiate namespace.

