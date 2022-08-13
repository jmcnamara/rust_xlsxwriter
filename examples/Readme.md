# Examples for the rust_xlsxwriter library.

This directory contains working examples showing different features of the
rust_xlsxwriter library.

The `app_{name}.rs` examples are small complete programs showing a feature or
collection of features. The `doc_{struct}_{function}.rs` examples are more
specific and generally show how an individual function works. The `doc_*.rs`
examples are usually repeated in the documentation.

* app_demo.rs - A simple, getting started, example of some of the features
  of the rust_xlsxwriter library.

* app_perf_test.rs - Simple performance test for rust_xlsxwriter.

* doc_worksheet_set_name.rs - Demonstrates setting user defined worksheet
  names and the default values when a name isn't set.

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

