# Examples for the rust_xlsxwriter library.

This directory contains working examples showing different features of the
rust_xlsxwriter library.

The `app_{name}.rs` examples are small complete programs showing a feature or
collection of features.

The `doc_{struct}_{function}.rs` examples are more specific examples from the
documentation and generally show how an individual function works.

* app_colors.rs - A sample palette of the the defined colors and user
  defined RGB colors available in the rust_xlsxwriter library.

* app_demo.rs - A simple, getting started, example of some of the features
  of the rust_xlsxwriter library.

* app_formatting.rs - An example of the various cell formatting options
  that are available in the rust_xlsxwriter library. These are laid out on
  worksheets that correspond to the sections of the Excel "Format Cells"
  dialog.

* app_hello_world.rs - Create a simple Hello World style Excel spreadsheet
  using the rust_xlsxwriter library.

* app_perf_test.rs - Simple performance test for rust_xlsxwriter.

* app_right_to_left.rs - Example of using rust_xlsxwriter to create a
  workbook with the default worksheet and cell text direction changed from
  left-to-right to right-to-left, as required by some middle eastern
  versions of Excel.

* app_tutorial1.rs - A simple program to write some data to an Excel
  spreadsheet using rust_xlsxwriter. Part 1 of a tutorial.

* app_tutorial2.rs - A simple program to write some data to an Excel
  spreadsheet using rust_xlsxwriter. Part 2 of a tutorial.

* app_tutorial3.rs - A simple program to write some data to an Excel
  spreadsheet using rust_xlsxwriter. Part 3 of a tutorial.

* doc_enum_xlsxcolor.rs - Demonstrates using different XlsxColor enum
  values to set the color of some text in a worksheet.

* doc_format_clone.rs - Demonstrates cloning a format and setting the
  properties.

* doc_format_create.rs - Demonstrates create a new format and setting the
  properties.

* doc_format_currency1.rs - Demonstrates setting a currency format for a
  worksheet cell. This example doesn't actually set a currency format, for
  that see the followup example in doc_format_currency2.rs.

* doc_format_currency2.rs - Demonstrates setting a currency format for a
  worksheet cell.

* doc_format_default.rs - Demonstrates creating a default format.

* doc_format_intro.rs - Demonstrates some of the available formatting
  properties.

* doc_format_locale.rs - Demonstrates setting a number format that appears
  differently in different locales.

* doc_format_new.rs - Demonstrates creating a new format.

* doc_format_set_align.rs - Demonstrates setting various cell alignment
  properties.

* doc_format_set_background_color.rs - Demonstrates setting the cell
  background color, with a default solid pattern.

* doc_format_set_bold.rs - Demonstrates setting the bold property for a
  format.

* doc_format_set_border.rs - Demonstrates setting a cell border.

* doc_format_set_border_color.rs - Demonstrates setting a cell border and
  color.

* doc_format_set_border_diagonal.rs - Demonstrates setting cell diagonal
  borders.

* doc_format_set_font_color.rs - Demonstrates setting the italic property
  for a format.

* doc_format_set_font_name.rs - Demonstrates setting the font name/type for
  a format.

* doc_format_set_font_size.rs - Demonstrates setting the font size for a
  format.

* doc_format_set_font_strikethrough.rs - Demonstrates setting the text
  strikethrough property for a format.

* doc_format_set_foreground_color.rs - Demonstrates setting the
  foreground/pattern color.

* doc_format_set_indent.rs - Demonstrates setting the indentation level for
  cell text.

* doc_format_set_italic.rs - Demonstrates setting the italic property for a
  format.

* doc_format_set_num_format.rs - Demonstrates setting different types of
  Excel number formatting.

* doc_format_set_num_format_index.rs - Demonstrates setting one of the
  inbuilt format indices for a format.

* doc_format_set_pattern.rs - Demonstrates setting the cell pattern (with
  colors).

* doc_format_set_reading_direction.rs - Demonstrates setting the text
  reading direction. This is useful when creating Arabic, Hebrew or other
  near or far eastern worksheets.

* doc_format_set_rotation.rs - Demonstrates setting text rotation for a
  cell.

* doc_format_set_shrink.rs - Demonstrates setting the text shrink format.

* doc_format_set_text_wrap.rs - Demonstrates setting an implicit (without
  newline) text wrap and a user defined text wrap (with newlines).

* doc_format_set_underline.rs - Demonstrates setting underline properties
  for a format.

* doc_worksheet_set_column_format.rs - Demonstrates setting the format for
  a column in Excel.

* doc_worksheet_set_column_width.rs - Demonstrates setting the width of
  columns in Excel.

* doc_worksheet_set_column_width_pixels.rs - Demonstrates setting the width
  of columns in Excel in pixels.

* doc_worksheet_set_formula_result.rs - Demonstrates manually setting the
  result of a formula. Note, this is only required for non-Excel
  applications that don't calculate formula results.

* doc_worksheet_set_formula_result_default.rs - Demonstrates manually
  setting the default result for all non-calculated formulas in a
  worksheet.

* doc_worksheet_set_name.rs - Demonstrates setting user defined worksheet
  names and the default values when a name isn't set.

* doc_worksheet_set_paper.rs - Demonstrates setting the worksheet paper
  size/type for the printed output.

* doc_worksheet_set_right_to_left.rs - Demonstrates changing the default
  worksheet and cell text direction changed from left-to-right to
  right-to-left, as required by some middle eastern versions of Excel.

* doc_worksheet_set_row_format.rs - Demonstrates setting the format for a
  row in Excel.

* doc_worksheet_set_row_height.rs - Demonstrates setting the height for a
  row in Excel.

* doc_worksheet_set_row_height_pixels.rs - Demonstrates setting the height
  for a row in Excel.

* doc_worksheet_write_blank.rs - Demonstrates writing a blank cell with
  formatting, i.e., a cell that has no data but does have formatting.

* doc_worksheet_write_date.rs - Demonstrates writing formatted dates in an
  Excel worksheet.

* doc_worksheet_write_datetime.rs - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* doc_worksheet_write_formula.rs - Demonstrates writing formulas with
  formatting to a worksheet.

* doc_worksheet_write_formula_only.rs - Demonstrates writing formulas with
  formatting to a worksheet.

* doc_worksheet_write_number.rs - Demonstrates setting different formatting
  for numbers in an Excel worksheet.

* doc_worksheet_write_number_only.rs - Demonstrates writing unformatted
  numbers to an Excel worksheet. Any numeric type that will convert
  [`Into`] f64 can be transferred to Excel.

* doc_worksheet_write_string.rs - Demonstrates setting different formatting
  for numbers in an Excel worksheet.

* doc_worksheet_write_string_only.rs - Demonstrates writing some UTF-8
  strings to a worksheet. The UTF-8 encoding is the only encoding supported
  by the Excel file format.

* doc_worksheet_write_time.rs - Demonstrates writing formatted times in an
  Excel worksheet.

