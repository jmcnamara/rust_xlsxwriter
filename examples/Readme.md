# Examples for the rust_xlsxwriter library.

This directory contains working examples showing different features of the
rust_xlsxwriter library.

The `app_{name}.rs` examples are small complete programs showing a feature or
collection of features.

The `doc_{struct}_{function}.rs` examples are more specific examples from the
documentation and generally show how an individual function works.

* app_array_formula.rs - Example of how to use the rust_xlsxwriter to write
  simple array formulas.

* app_autofit.rs - An example of using a simulated Column Autofit using the
  rust_xlsxwriter library.

* app_colors.rs - A demonstration of the RGB and Theme colors palettes
  available in the rust_xlsxwriter library.

* app_demo.rs - A simple, getting started, example of some of the features
  of the rust_xlsxwriter library.

* app_dynamic_arrays.rs - An example of how to use the rust_xlsxwriter
  library to write formulas and functions that create dynamic arrays. These
  functions are new to Excel 365. The examples mirror the examples in the
  Excel documentation for these functions.

* app_file_to_memory.rs - An example of creating a simple Excel xlsx file
  in an in memory Vec<u8> buffer using the rust_xlsxwriter library.

* app_formatting.rs - An example of the various cell formatting options
  that are available in the rust_xlsxwriter library. These are laid out on
  worksheets that correspond to the sections of the Excel "Format Cells"
  dialog.

* app_headers_footers.rs - An example of setting headers and footers in
  worksheets using the rust_xlsxwriter library.

* app_hello_world.rs - Create a simple Hello World style Excel spreadsheet
  using the rust_xlsxwriter library.

* app_hyperlinks.rs - A simple, getting started, example of some of the
  features of the rust_xlsxwriter library.

* app_images.rs - An example of inserting images into a worksheet using
  rust_xlsxwriter.

* app_merge_range.rs - An example of creating merged ranges in a worksheet
  using the rust_xlsxwriter library.

* app_panes.rs - A simple example of setting some "freeze" panes in
  worksheets using the rust_xlsxwriter library.

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

* doc_image.rs - This example shows how to create an image object and use
  it to insert the image into a worksheet.

* doc_image_dimensions.rs - This example shows how to get some of the
  properties of an Image that will be used in an Excel worksheet.

* doc_image_set_scale_width.rs - This example shows how to create an image
  object and use it to insert the image into a worksheet. The image in this
  case is scaled.

* doc_workbook_add_worksheet.rs - Demonstrates creating adding worksheets
  to a workbook.

* doc_workbook_new.rs - Demonstrates creating a simple workbook, with one
  unused worksheet.

* doc_workbook_push_worksheet.rs - Demonstrates creating a standalone
  worksheet object and then adding it to a workbook.

* doc_workbook_save.rs - Demonstrates creating a simple workbook, with one
  unused worksheet.

* doc_workbook_save_to_buffer.rs - Demonstrates creating a simple workbook
  to a Vec<u8> buffer.

* doc_workbook_save_to_path.rs - Demonstrates creating a simple workbook
  using a rust Path reference.

* doc_workbook_worksheet_from_index.rs - Demonstrates getting worksheet
  reference by index.

* doc_workbook_worksheet_from_name.rs - Demonstrates getting worksheet
  reference by name.

* doc_workbook_worksheets.rs - Demonstrates operating on the vector of all
  the worksheets in a workbook. The non mutable version of this method is
  less useful than `workbook.worksheets_mut()`.

* doc_workbook_worksheets_mut.rs - Demonstrates operating on the vector of
  all the worksheets in a workbook.

* doc_working_with_formulas_dynamic_len.rs - Demonstrates a static function
  which generally returns one value turned into a dynamic function which
  returns a range of values.

* doc_working_with_formulas_future1.rs - Demonstrates writing an Excel
  "Future Function" without an explicit prefix, which results in an Excel
  error.

* doc_working_with_formulas_future2.rs - Demonstrates writing an Excel
  "Future Function" with an explicit prefix.

* doc_working_with_formulas_future3.rs - Demonstrates writing an Excel
  "Future Function" with an implicit prefix and the use_future_functions()
  method.

* doc_working_with_formulas_intro.rs - Demonstrates a simple formula.

* doc_working_with_formulas_static_len.rs - Demonstrates a static function
  which generally returns one value. Compare this with the dynamic function
  output of doc_working_with_formulas_dynamic_len.rs.

* doc_working_with_formulas_syntax.rs - Demonstrates some common formula
  syntax errors.

* doc_worksheet_insert_image_with_offset.rs - This example shows how to add
  an image to a worksheet at an offset within the cell.

* doc_worksheet_name.rs - Demonstrates getting a worksheet name.

* doc_worksheet_new.rs - Demonstrates creating new worksheet objects and
  then adding them to a workbook.

* doc_worksheet_set_active.rs - Demonstrates setting a worksheet as the
  visible worksheet when a file is opened.

* doc_worksheet_set_autofit.rs - Demonstrates auto-fitting the worksheet
  column widths based on the data in the columns.

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

* doc_worksheet_set_freeze_panes.rs - Demonstrates setting the worksheet
  panes.

* doc_worksheet_set_freeze_panes_top_cell.rs - Demonstrates setting the
  worksheet panes and also setting the topmost visible cell in the scrolled
  area.

* doc_worksheet_set_header.rs - Demonstrates setting the worksheet header.

* doc_worksheet_set_hidden.rs - Demonstrates hiding a worksheet.

* doc_worksheet_set_landscape.rs - Demonstrates setting the worksheet page
  orientation to landscape.

* doc_worksheet_set_margins.rs - Demonstrates setting the worksheet
  margins.

* doc_worksheet_set_name.rs - Demonstrates setting user defined worksheet
  names and the default values when a name isn't set.

* doc_worksheet_set_page_order.rs - Demonstrates setting the worksheet
  printed page order.

* doc_worksheet_set_paper.rs - Demonstrates setting the worksheet paper
  size/type for the printed output.

* doc_worksheet_set_print_area.rs - Demonstrates setting the print area for
  several worksheets.

* doc_worksheet_set_print_first_page_number.rs - Demonstrates setting the
  page number on the printed page.

* doc_worksheet_set_print_fit_to_pages.rs - Demonstrates setting the scale
  of the worksheet to fit a defined number of pages vertically and
  horizontally. This example shows a common use case which is to fit the
  printed output to 1 page wide but have the height be as long as
  necessary.

* doc_worksheet_set_print_scale.rs - Demonstrates setting the scale of the
  worksheet page when printed.

* doc_worksheet_set_repeat_columns.rs - Demonstrates setting the columns to
  repeat on each printed page.

* doc_worksheet_set_repeat_rows.rs - Demonstrates setting the rows to
  repeat on each printed page.

* doc_worksheet_set_right_to_left.rs - Demonstrates changing the default
  worksheet and cell text direction changed from left-to-right to
  right-to-left, as required by some middle eastern versions of Excel.

* doc_worksheet_set_row_format.rs - Demonstrates setting the format for a
  row in Excel.

* doc_worksheet_set_row_height.rs - Demonstrates setting the height for a
  row in Excel.

* doc_worksheet_set_row_height_pixels.rs - Demonstrates setting the height
  for a row in Excel.

* doc_worksheet_set_selected.rs - Demonstrates selecting worksheet in a
  workbook. The active worksheet is selected by default so in this example
  the first two worksheets are selected.

* doc_worksheet_set_tab_color.rs - Demonstrates set the tab color of
  worksheets.

* doc_worksheet_set_zoom.rs - Demonstrates setting the worksheet zoom
  level.

* doc_worksheet_write_array_formula.rs - Demonstrates writing an array
  formulas with formatting to a worksheet.

* doc_worksheet_write_array_formula_only.rs - Demonstrates writing an array
  formulas to a worksheet.

* doc_worksheet_write_blank.rs - Demonstrates writing a blank cell with
  formatting, i.e., a cell that has no data but does have formatting.

* doc_worksheet_write_boolean.rs - Demonstrates writing formatted boolean
  values to a worksheet.

* doc_worksheet_write_boolean_only.rs - Demonstrates writing boolean values
  to a worksheet.

* doc_worksheet_write_date.rs - Demonstrates writing formatted dates in an
  Excel worksheet.

* doc_worksheet_write_datetime.rs - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* doc_worksheet_write_dynamic_array_formula.rs - Demonstrates a static
  function which generally returns one value turned into a dynamic array
  function which returns a range of values.

* doc_worksheet_write_dynamic_array_formula_only.rs - Demonstrates a static
  function which generally returns one value turned into a dynamic array
  function which returns a range of values.

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

* doc_worksheet_write_url_with_format.rs - Demonstrates writing a url with
  alternative format.

* doc_worksheet_write_url_with_text.rs - Demonstrates writing a url with
  alternative text.

