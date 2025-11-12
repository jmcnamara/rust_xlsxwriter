# Examples for the `rust_xlsxwriter` library.

This directory contains working examples showing different features of the
`rust_xlsxwriter` library.

The `app_{name}.rs` examples are small complete programs showing a feature or
collection of features.

The `doc_{struct}_{function}.rs` examples are more specific examples from the
documentation and generally show how an individual function works.

* `app_array_formula.rs` - Example of how to use the `rust_xlsxwriter` to
  write simple array formulas.

* `app_autofilter.rs` - An example of how to create autofilters with the
  `rust_xlsxwriter` library. An autofilter is a way of adding drop down
  lists to the headers of a 2D range of worksheet data. This allows users
  to filter the data based on simple criteria so that some data is shown
  and some is hidden.

* `app_autofit.rs` - An example of using simulated autofit to automatically
  adjust the width of worksheet columns based on the data in the cells.

* `app_background_image.rs` - An example of inserting a background image
  into a worksheet using `rust_xlsxwriter`. See also the `app_watermark.rs`
  example which shows how to set a watermark via the header image of a
  worksheet. That is the way that the Microsoft documentation recommends to
  set a watermark in Excel.

* `app_chart_area.rs` - An example of creating area charts using the
  `rust_xlsxwriter` library.

* `app_chart_bar.rs` - An example of creating bar charts using the
  `rust_xlsxwriter` library.

* `app_chart_clustered.rs` - A demo of a clustered category chart using
  rust_xlsxwriter.

* `app_chart_column.rs` - An example of creating column charts using the
  `rust_xlsxwriter` library.

* `app_chart_combined.rs` - An example of creating combined charts using
  the `rust_xlsxwriter` library.

* `app_chart_data_table.rs` - An example of creating Excel Column charts
  with data tables using the `rust_xlsxwriter` library.

* `app_chart_data_tools.rs` - A demo of the various Excel chart data tools
  that are available via the `rust_xlsxwriter` library.

* `app_chart_doughnut.rs` - An example of creating doughnut charts using
  the `rust_xlsxwriter` library.

* `app_chart_gauge.rs` - An example of creating a Gauge Chart in Excel
  using the `rust_xlsxwriter` library. A Gauge Chart isn't a native chart
  type in Excel. It is constructed by combining a doughnut chart and a pie
  chart and by using some non-filled elements to hide parts of the default
  charts. This example follows the following online example of how to
  create a [Gauge Chart] in Excel. [Gauge Chart]:
  https://www.excel-easy.com/examples/gauge-chart.html

* `app_chart_gradient.rs` - An example of creating a chart with gradient
  fills using the `rust_xlsxwriter` library.

* `app_chart_line.rs` - An example of creating line charts using the
  `rust_xlsxwriter` library.

* `app_chart_pareto.rs` - An example of creating a Pareto chart using the
  `rust_xlsxwriter` library.

* `app_chart_pattern.rs` - An example of creating column charts with fill
  patterns using the `rust_xlsxwriter` library.

* `app_chart_pie.rs` - An example of creating pie charts using the
  `rust_xlsxwriter` library.

* `app_chart_radar.rs` - An example of creating radar charts using the
  `rust_xlsxwriter` library.

* `app_chart_scatter.rs` - An example of creating scatter charts using the
  `rust_xlsxwriter` library.

* `app_chart_secondary_axis.rs` - An example of creating an Excel Line
  chart with a secondary axis using the `rust_xlsxwriter` library.

* `app_chart_stock.rs` - An example of creating Stock charts using the
  `rust_xlsxwriter` library. Note, Volume variants of the Excel stock
  charts aren't currently supported but will be in a future release.

* `app_chart_styles.rs` - An example showing all 48 default chart styles
  available in Excel 2007 using `rust_xlsxwriter`. Note, these styles are
  not the same as the styles available in Excel 2013 and later.

* `app_chart_tutorial1.rs` - An example of creating a simple chart using
  the `rust_xlsxwriter` library.

* `app_chart_tutorial2.rs` - An example of creating a simple chart using
  the `rust_xlsxwriter` library.

* `app_chart_tutorial3.rs` - An example of creating a simple chart using
  the `rust_xlsxwriter` library.

* `app_chart_tutorial4.rs` - An example of creating a simple chart using
  the `rust_xlsxwriter` library.

* `app_chart.rs` - A simple chart example using the `rust_xlsxwriter`
  library.

* `app_chartsheet.rs` - An example of creating a chartsheet style chart
  using the `rust_xlsxwriter` library.

* `app_checkbox.rs` - An example of adding checkbox boolean values to a
  worksheet using the `rust_xlsxwriter` library.

* `app_colors.rs` - A demonstration of the RGB and Theme colors palettes
  available in the `rust_xlsxwriter` library.

* `app_conditional_formatting.rs` - Example of how to add conditional
  formatting to a worksheet using the `rust_xlsxwriter` library.
  Conditional formatting allows you to apply a format to a cell or a range
  of cells based on user defined rule.

* `app_data_validation.rs` - Example of how to add data validation and
  dropdown lists using the `rust_xlsxwriter` library. Data validation is a
  feature of Excel which allows you to restrict the data that a user enters
  in a cell and to display help and warning messages. It also allows you to
  restrict input to values in a drop down list.

* `app_defined_name.rs` - Example of how to create defined names using the
  `rust_xlsxwriter` library. This functionality is used to define user
  friendly variable names to represent a value, a single cell,	or a range
  of cells in a workbook.

* `app_demo.rs` - A simple, getting started, example of some of the
  features of the `rust_xlsxwriter` library.

* `app_doc_properties.rs` - An example of setting workbook document
  properties for a file created using the `rust_xlsxwriter` library.

* `app_dynamic_arrays.rs` - An example of how to use the `rust_xlsxwriter`
  library to write formulas and functions that create dynamic arrays. These
  functions are new to Excel 365. The examples mirror the examples in the
  Excel documentation for these functions.

* `app_embedded_images.rs` - An example of embedding images into a
  worksheet cells using `rust_xlsxwriter`. This image scales to size of the
  cell and moves with it. This approach can be useful if you are building
  up a spreadsheet of products with a column of images for each product.
  This is the equivalent of Excel's menu option to insert an image using
  the option to "Place in Cell" which is only available in Excel 365
  versions from 2023 onwards. For older versions of Excel a `#VALUE!` error
  is displayed.

* `app_file_to_memory.rs` - An example of creating a simple Excel xlsx file
  in an in memory Vec<u8> buffer using the `rust_xlsxwriter` library.

* `app_formatting.rs` - An example of the various cell formatting options
  that are available in the `rust_xlsxwriter` library. These are laid out
  on worksheets that correspond to the sections of the Excel "Format Cells"
  dialog.

* `app_grouped_columns.rs` - An example of how to group columns into
  outlines with the `rust_xlsxwriter` library. In Excel an outline is a
  group of rows or columns that can be collapsed or expanded to simplify
  hierarchical data. It is often used with the `SUBTOTAL()` function.

* `app_grouped_rows.rs` - An example of how to group rows into outlines
  with the `rust_xlsxwriter` library. In Excel an outline is a group of
  rows or columns that can be collapsed or expanded to simplify
  hierarchical data. It is often used with the `SUBTOTAL()` function.

* `app_headers_footers.rs` - An example of setting headers and footers in
  worksheets using the `rust_xlsxwriter` library.

* `app_hello_world.rs` - Create a simple Hello World style Excel
  spreadsheet using the `rust_xlsxwriter` library.

* `app_hyperlinks.rs` - An example of some of the features of
  URLs/hyperlinks using the `rust_xlsxwriter` library.

* `app_ignore_errors.rs` - An example of turning off worksheet cells
  errors/warnings using using the `rust_xlsxwriter` library.

* `app_images_fit_to_cell.rs` - An example of inserting images into a
  worksheet using `rust_xlsxwriter` so that they are scaled to a cell. This
  approach can be useful if you are building a spreadsheet of products with
  a column of images for each product. See also the `app_embedded_image.rs`
  example that shows a better approach for newer versions of Excel.

* `app_images.rs` - An example of inserting images into a worksheet using
  `rust_xlsxwriter`.

* `app_lambda.rs` - An example of using the new Excel LAMBDA() function
  with the `rust_xlsxwriter` library.

* `app_macros.rs` - An example of adding macros to an `rust_xlsxwriter`
  file using a VBA macros file extracted from an existing Excel xlsm file.
  The `vba_extract` utility (https://crates.io/crates/vba_extract) can be
  used to extract the `vbaProject.bin` file.

* `app_memory_test.rs` - Simple performance test and memory usage program
  for `rust_xlsxwriter`. It writes alternate cells of strings and numbers.
  It defaults to 4,000 rows x 40 columns. The number of rows and the
  "constant memory" mode can be optionally set. usage:
  ./target/release/examples/app_perf_test [num_rows] [--constant-memory]

* `app_merge_range.rs` - An example of creating merged ranges in a
  worksheet using the `rust_xlsxwriter` library.

* `app_notes.rs` - An example of writing cell Notes to a worksheet using
  the `rust_xlsxwriter` library.

* `app_panes.rs` - A simple example of setting some "freeze" panes in
  worksheets using the `rust_xlsxwriter` library.

* `app_perf_test.rs` - Simple performance test program for
  `rust_xlsxwriter`. It writes alternate cells of strings and numbers. It
  defaults to 4,000 rows x 40 columns. usage:
  ./target/release/examples/app_perf_test [num_rows]

* `app_rich_strings.rs` - An example of using the `rust_xlsxwriter` library
  to write "rich" multi-format strings in worksheet cells.

* `app_right_to_left.rs` - Example of using `rust_xlsxwriter` to create a
  workbook with the default worksheet and cell text direction changed from
  left-to-right to right-to-left, as required by some middle eastern
  versions of Excel.

* `app_sensitivity_label.rs` - An example adding a Sensitivity Label to an
  Excel file using custom document properties. See the main docs for an
  explanation of how to extract the metadata.

* `app_serialize.rs` - Example of serializing Serde derived structs to an
  Excel worksheet using `rust_xlsxwriter`.

* `app_sparklines1.rs` - Example of adding sparklines to an Excel
  spreadsheet using the `rust_xlsxwriter` library. Sparklines are small
  charts that fit in a single cell and are used to show trends in data.

* `app_sparklines2.rs` - Example of adding sparklines to an Excel
  spreadsheet using the `rust_xlsxwriter` library. Sparklines are small
  charts that fit in a single cell and are used to show trends in data.
  This example shows the majority of the properties that can applied to
  sparklines.

* `app_table_of_contents.rs` - This is an example of creating a "Table of
  Contents" worksheet with links to other worksheets in the workbook.

* `app_tables.rs` - Example of how to add tables to a worksheet using the
  `rust_xlsxwriter` library. Tables in Excel are used to group rows and
  columns of data into a single structure that can be referenced in a
  formula or formatted collectively.

* `app_textbox.rs` - Demonstrate adding a Textbox to a worksheet using the
  `rust_xlsxwriter` library.

* `app_theme_custom.rs` - Example of setting the default theme for a
  workbook to a user supplied custom theme using the `rust_xlsxwriter`
  library. The theme xml file is extracted from an Excel xlsx file.

* `app_theme_excel_2023.rs` - Example of changing the default theme for a
  workbook using the `rust_xlsxwriter` library. The example uses the Excel
  2023 Office/Aptos theme.

* `app_tutorial1.rs` - A simple program to write some data to an Excel
  spreadsheet using `rust_xlsxwriter`. Part 1 of a tutorial.

* `app_tutorial2.rs` - A simple program to write some data to an Excel
  spreadsheet using `rust_xlsxwriter`. Part 2 of a tutorial.

* `app_tutorial3.rs` - A simple program to write some data to an Excel
  spreadsheet using `rust_xlsxwriter`. Part 3 of a tutorial.

* `app_tutorial4.rs` - A simple program to write some data to an Excel
  spreadsheet using `rust_xlsxwriter`. Part 4 of a tutorial.

* `app_tutorial5.rs` - A simple program to write some data to an Excel
  spreadsheet using `rust_xlsxwriter`. Part 5 of a tutorial.

* `app_watermark.rs` - An example of adding a worksheet watermark image
  using the `rust_xlsxwriter` library. This is based on the method of
  putting an image in the worksheet header as suggested in the Microsoft
  documentation.

* `app_worksheet_protection.rs` - Example of cell locking and formula
  hiding in an Excel worksheet `rust_xlsxwriter` library.

* `app_write_arrays.rs` - An example of writing arrays of data using the
  `rust_xlsxwriter` library. Array in this context means Rust arrays or
  arrays like data types that implement `IntoIterator`. The array must also
  contain data types that implement `rust_xlsxwriter`'s `IntoExcelData`.

* `app_write_generic_data.rs` - Example of how to extend the the
  `rust_xlsxwriter` `write()` method using the IntoExcelData trait to
  handle arbitrary user data that can be mapped to one of the main Excel
  data types.

* `doc_button_set_caption.rs` - An example of adding an Excel Form Control
  button to a worksheet. This example demonstrates setting the button
  caption.

* `doc_button_set_macro.rs` - An example of adding an Excel Form Control
  button to a worksheet. This example demonstrates setting the button
  macro.

* `doc_chart_add_series.rs` - An example of creating a chart series via
  [`Chart::add_series()`](Chart::add_series).

* `doc_chart_axis_set_crossing.rs` - A chart example demonstrating setting
  the point where the axes will cross.

* `doc_chart_axis_set_date_axis.rs` - A chart example demonstrating setting
  a date axis for a chart.

* `doc_chart_axis_set_display_unit_type.rs` - A chart example demonstrating
  setting the units of the Value/Y-axis.

* `doc_chart_axis_set_hidden.rs` - A chart example demonstrating hiding the
  chart axes.

* `doc_chart_axis_set_label_interval.rs` - A chart example demonstrating
  setting the label interval for an axis.

* `doc_chart_axis_set_label_position.rs` - A chart example demonstrating
  setting the label position for an axis.

* `doc_chart_axis_set_log_base.rs` - A chart example demonstrating setting
  the logarithm base for chart axes.

* `doc_chart_axis_set_major_gridlines_line.rs` - A chart example
  demonstrating formatting the major gridlines for chart axes.

* `doc_chart_axis_set_major_gridlines.rs` - A chart example demonstrating
  turning off the major gridlines for chart axes.

* `doc_chart_axis_set_major_tick_type.rs` - A chart example demonstrating
  setting the tick types for chart axes.

* `doc_chart_axis_set_major_unit.rs` - A chart example demonstrating
  setting the units for chart axes.

* `doc_chart_axis_set_max_date.rs` - A chart example demonstrating setting
  the maximum and minimum values for a date axis.

* `doc_chart_axis_set_max.rs` - A chart example demonstrating setting the
  axes bounds for chart axes.

* `doc_chart_axis_set_minor_gridlines_line.rs` - A chart example
  demonstrating formatting the minor gridlines for chart axes.

* `doc_chart_axis_set_minor_gridlines.rs` - A chart example demonstrating
  turning on the minor gridlines for chart axes.

* `doc_chart_axis_set_name_font.rs` - An example of setting the font for a
  chart axis title.

* `doc_chart_axis_set_name_format.rs` - A chart example demonstrating
  setting the formatting of the title of chart axes.

* `doc_chart_axis_set_name.rs` - A chart example demonstrating setting the
  title of chart axes.

* `doc_chart_axis_set_num_format.rs` - A chart example demonstrating
  setting the number format a chart axes.

* `doc_chart_axis_set_position_between_ticks.rs` - A chart example
  demonstrating setting the axes data position relative to the tick marks.
  Notice that by setting the data columns "on" the tick the first and last
  columns are cut off by the plot area.

* `doc_chart_axis_set_reverse.rs` - A chart example demonstrating reversing
  the plotting direction of the chart axes.

* `doc_chart_axis_set_tick_interval.rs` - A chart example demonstrating
  setting the tick interval for an axis.

* `doc_chart_border_formatting.rs` - An example of formatting the border in
  a chart element.

* `doc_chart_combine1.rs` - An example of creating a combined Column and
  Line chart. In this example they share the same primary Y axis.

* `doc_chart_combine2.rs` - An example of creating a combined Column and
  Line chart. In this example the Column values are on the primary Y axis
  and the Line chart values are on the secondary Y2 axis.

* `doc_chart_data_labels_set_font.rs` - An example of adding data labels to
  a chart series with font formatting.

* `doc_chart_data_labels_set_format.rs` - An example of adding data labels
  to a chart series with formatting.

* `doc_chart_data_labels_set_num_format.rs` - An example of adding data
  labels to a chart series with number formatting.

* `doc_chart_data_labels_set_position.rs` - An example of adding data
  labels to a chart series and changing their default position.

* `doc_chart_data_labels_show_category_name.rs` - An example of adding data
  labels to a chart series with value and category details.

* `doc_chart_data_labels_show_percentage.rs` - An example of setting the
  percentage for the data labels of a chart series. Usually this only
  applies to a Pie or Doughnut chart.

* `doc_chart_data_labels.rs` - An example of adding data labels to a chart
  series.

* `doc_chart_error_bars_intro.rs` - An example of adding error bars to a
  chart data series.

* `doc_chart_font_set_bold.rs` - An example of setting the bold property
  for the font in a chart element.

* `doc_chart_font_set_color.rs` - An example of setting the color property
  for the font in a chart element.

* `doc_chart_font_set_italic.rs` - An example of setting the italic
  property for the font in a chart element.

* `doc_chart_font_set_name.rs` - An example of setting the font name
  property for the font in a chart element.

* `doc_chart_font_set_rotation.rs` - An example of setting the font text
  rotation for the font in a chart element.

* `doc_chart_font_set_size.rs` - An example of setting the font size
  property for the font in a chart element.

* `doc_chart_font.rs` - An example of setting the font for a chart element.

* `doc_chart_format_set_gradient_fill.rs` - An example of setting a
  gradient fill for a chart element.

* `doc_chart_format_set_no_border.rs` - An example of turning off the
  border of a chart element.

* `doc_chart_format_set_no_fill.rs` - An example of turning off the fill of
  a chart element.

* `doc_chart_format_set_no_line.rs` - An example of turning off a default
  line in a chart format.

* `doc_chart_formatting.rs` - An example of formatting the chart border
  element.

* `doc_chart_gradient_fill_set_type.rs` - An example of setting a gradient
  fill for a chart element with a non-default gradient type.

* `doc_chart_gradient_fill.rs` - An example of setting a gradient fill for
  a chart element.

* `doc_chart_gradient_stops_new.rs` - An example of creating gradient stops
  for a gradient fill for a chart element.

* `doc_chart_gradient_stops.rs` - An example of setting a gradient fill for
  a chart element.

* `doc_chart_gradient_stops2.rs` - An example of setting a gradient fill
  for a chart element.

* `doc_chart_intro.rs` - A simple chart example using the `rust_xlsxwriter`
  library.

* `doc_chart_legend_delete_entries.rs` - A chart example demonstrating
  deleting/hiding a series name from the chart legend.

* `doc_chart_legend_set_font.rs` - An example of setting the font for a
  chart legend.

* `doc_chart_legend_set_hidden.rs` - An example of hiding a default chart
  legend.

* `doc_chart_legend_set_overlay.rs` - An example of overlaying the chart
  legend on the plot area.

* `doc_chart_legend.rs` - An example of getting the chart legend object and
  setting some of its properties.

* `doc_chart_line_formatting.rs` - An example of formatting a line/border
  in a chart element.

* `doc_chart_line_set_color.rs` - An example of formatting the line color
  in a chart element.

* `doc_chart_line_set_dash_type.rs` - An example of formatting the line
  dash type in a chart element.

* `doc_chart_line_set_transparency.rs` - An example of formatting the line
  transparency in a chart element. Note, you must set also set a color in
  order to set the transparency.

* `doc_chart_line_set_width.rs` - An example of formatting the line width
  in a chart element.

* `doc_chart_marker_set_automatic.rs` - An example of adding automatic
  markers to a line chart.

* `doc_chart_marker_set_size.rs` - An example of adding markers to a line
  chart with user defined size.

* `doc_chart_marker_set_type.rs` - An example of adding markers to a line
  chart with user defined marker types.

* `doc_chart_marker.rs` - An example of adding markers to a line chart.

* `doc_chart_pattern_fill_set_pattern.rs` - An example of setting a pattern
  fill for a chart element.

* `doc_chart_pattern_fill.rs` - An example of setting a pattern fill for a
  chart element.

* `doc_chart_plot_area_set_layout.rs` - An example of setting the layout of
  a chart element, in this case the chart plot area.

* `doc_chart_push_series.rs` - An example of creating a chart series as a
  standalone object and then adding it to a chart via the
  [`Chart::push_series()`](Chart::add_series) method.

* `doc_chart_series_delete_from_legend.rs` - A chart example demonstrating
  deleting/hiding a series name from the chart legend.

* `doc_chart_series_set_categories.rs` - A chart example demonstrating
  setting the chart series categories and values.

* `doc_chart_series_set_invert_if_negative_color.rs` - A chart example
  demonstrating setting the "Invert if negative" property and associated
  color for a chart series. This also requires that you set a solid fill
  color for the series.

* `doc_chart_series_set_invert_if_negative.rs` - A chart example
  demonstrating setting the "Invert if negative" property for a chart
  series.

* `doc_chart_series_set_name.rs` - A chart example demonstrating setting
  the chart series name.

* `doc_chart_series_set_overlap.rs` - An example of setting the chart
  series gap and overlap. Note that it only needs to be applied to one of
  the series in the chart.

* `doc_chart_series_set_secondary_axis.rs` - A chart example demonstrating
  setting a secondary Y axis.

* `doc_chart_series_set_secondary_axis2.rs` - A chart example demonstrating
  using a secondary X and Y axis. The secondary X axis is only available
  for chart series that have a category range that is different from the
  primary category range.

* `doc_chart_series_set_values.rs` - A chart example demonstrating setting
  the chart series values.

* `doc_chart_set_chart_area_format.rs` - An example of formatting the chart
  "area" of a chart. In Excel the chart area is the background area behind
  the chart.

* `doc_chart_set_custom_data_labels1.rs` - An example of adding custom data
  labels to a chart series. This is useful when you want to label the
  points of a data series with information that isn't contained in the
  value or category names.

* `doc_chart_set_custom_data_labels2.rs` - An example of adding custom data
  labels to a chart series. This example shows how to get the data from
  cells. In Excel this is a single command called "Value from Cells" but in
  `rust_xlsxwriter` it needs to be broken down into a cell reference for
  each data label.

* `doc_chart_set_custom_data_labels3.rs` - An example of adding custom data
  labels to a chart series. This example shows how to add
  default/non-custom data labels along with custom data labels. This is
  done in two ways: with an explicit `default()` data label and with an
  implicit default for points that aren't covered at the end of the list.

* `doc_chart_set_custom_data_labels4.rs` - An example of adding custom data
  labels to a chart series. This example shows how to hide some of the data
  labels and keep others visible.

* `doc_chart_set_custom_data_labels5.rs` - An example of adding custom data
  labels to a chart series. This example shows how to format some of the
  data labels and leave the rest with the default formatting.

* `doc_chart_set_data_table.rs` - An example of adding a data table to a
  chart.

* `doc_chart_set_drop_lines_format.rs` - An example of setting drop lines
  for a chart, with formatting.

* `doc_chart_set_drop_lines.rs` - An example of setting drop lines for a
  chart.

* `doc_chart_set_high_low_lines_format.rs` - An example of setting high-low
  lines for a chart, with formatting.

* `doc_chart_set_high_low_lines.rs` - An example of setting high-low lines
  for a chart.

* `doc_chart_set_hole_size.rs` - An example of formatting the chart hole
  size for doughnut charts.

* `doc_chart_set_plot_area_format.rs` - An example of formatting the chart
  plot area of a chart. In Excel the plot area is the area between the axes
  on which the chart series are plotted.

* `doc_chart_set_point_colors.rs` - An example of setting the individual
  segment colors of a Pie chart.

* `doc_chart_set_points.rs` - An example of formatting the chart rotation
  for pie and doughnut charts.

* `doc_chart_set_rotation.rs` - An example of formatting the chart rotation
  for pie and doughnut charts.

* `doc_chart_set_up_down_bars_format.rs` - An example of setting up-down
  bars for a chart, with formatting.

* `doc_chart_set_up_down_bars.rs` - An example of setting up-down bars for
  a chart.

* `doc_chart_set_width.rs` - A simple chart example using the
  `rust_xlsxwriter` library.

* `doc_chart_simple.rs` - A simple chart example using the
  `rust_xlsxwriter` library.

* `doc_chart_solid_fill_set_color.rs` - An example of setting a solid fill
  color for a chart element.

* `doc_chart_solid_fill.rs` - An example of setting a solid fill for a
  chart element.

* `doc_chart_title_set_font.rs` - An example of setting the font for a
  chart title.

* `doc_chart_title_set_hidden.rs` - A simple chart example using the
  `rust_xlsxwriter` library.

* `doc_chart_title_set_name.rs` - A chart example demonstrating setting the
  chart title.

* `doc_chart_trendline_delete_from_legend.rs` - An example of adding a
  trendline to a chart data series. This demonstrates deleting/hiding the
  trendline name from the chart legend.

* `doc_chart_trendline_intro.rs` - An example of adding a trendline to a
  chart data series.

* `doc_chart_trendline_set_format.rs` - An example of adding a trendline to
  a chart data series with formatting.

* `doc_chart_trendline_set_label_format.rs` - An example of adding a
  trendline to a chart data series and adding formatting to the trendline
  data label.

* `doc_chart_trendline_set_name.rs` - An example of adding a trendline to a
  chart data series with a custom name.

* `doc_chart_trendline_set_type.rs` - An example of adding a trendline to a
  chart data series. Demonstrates setting the polynomial trendline type.

* `doc_chartrange_new_from_range.rs` - Demonstrates creating a new chart
  range.

* `doc_chartrange_new_from_string.rs` - Demonstrates creating a new chart
  range.

* `doc_chartsheet.rs` - A simple chartsheet example. A chart is placed on
  it own dedicated worksheet.

* `doc_conditional_format_2color_set_color.rs` - Example of adding a 2
  color scale type conditional formatting to a worksheet with user defined
  minimum and maximum colors.

* `doc_conditional_format_2color_set_minimum.rs` - Example of adding a 2
  color scale type conditional formatting to a worksheet with user defined
  minimum and maximum values.

* `doc_conditional_format_2color.rs` - Example of adding a 2 color scale
  type conditional formatting to a worksheet. Note, the colors in the fifth
  example (yellow to green) are the default colors and could be omitted.

* `doc_conditional_format_3color_set_color.rs` - Example of adding 3 color
  scale type conditional formatting to a worksheet with user defined
  minimum, midpoint and maximum colors.

* `doc_conditional_format_3color_set_minimum.rs` - Example of adding 3
  color scale type conditional formatting to a worksheet with user defined
  minimum and maximum values.

* `doc_conditional_format_3color.rs` - Example of adding 3 color scale type
  conditional formatting to a worksheet. Note, the colors in the first
  example (red to yellow to green) are the default colors and could be
  omitted.

* `doc_conditional_format_anchor.rs` - Example of adding a Formula type
  conditional formatting to a worksheet. This example demonstrate the
  effect of changing the absolute/relative anchor in the target cell.

* `doc_conditional_format_average.rs` - Example of how to add Average
  conditional formatting to a worksheet. Above average values are in light
  red. Below average values are in light green.

* `doc_conditional_format_blank.rs` - Example of how to add a
  blank/non-blank conditional formatting to a worksheet. Blank values are
  in light red. Non-blank values are in light green. Note, that we invert
  the Blank rule to get Non-blank values.

* `doc_conditional_format_cell_set_minimum.rs` - Example of adding a cell
  type conditional formatting to a worksheet. Values between 40 and 60 are
  highlighted in light green.

* `doc_conditional_format_cell_set_value.rs` - Example of adding a cell
  type conditional formatting to a worksheet. Cells with values >= 50 are
  in light green.

* `doc_conditional_format_cell1.rs` - Example of adding a cell type
  conditional formatting to a worksheet. Cells with values >= 50 are in
  light red. Values < 50 are in light green.

* `doc_conditional_format_cell2.rs` - Example of adding a cell type
  conditional formatting to a worksheet. Values between 30 and 70 are
  highlighted in light red. Values outside that range are in light green.

* `doc_conditional_format_databar_set_axis_color.rs` - Example of adding a
  data bar type conditional formatting to a worksheet with a user defined
  axis color.

* `doc_conditional_format_databar_set_axis_position.rs` - Example of adding
  a data bar type conditional formatting to a worksheet with different axis
  positions.

* `doc_conditional_format_databar_set_bar_only.rs` - Example of adding a
  data bar type conditional formatting to a worksheet with the bar only and
  with the data hidden.

* `doc_conditional_format_databar_set_border_color.rs` - Example of adding
  a data bar type conditional formatting to a worksheet with user defined
  border color.

* `doc_conditional_format_databar_set_border_off.rs` - Example of adding a
  data bar type conditional formatting to a worksheet without a border.

* `doc_conditional_format_databar_set_direction.rs` - Example of adding a
  data bar type conditional formatting to a worksheet without a border

* `doc_conditional_format_databar_set_fill_color.rs` - Example of adding a
  data bar type conditional formatting to a worksheet with user defined
  fill color.

* `doc_conditional_format_databar_set_minimum.rs` - Example of adding a
  data bar type conditional formatting to a worksheet with user defined
  minimum and maximum values.

* `doc_conditional_format_databar_set_negative_border_color.rs` - Example
  of adding a data bar type conditional formatting to a worksheet with user
  defined negative border color.

* `doc_conditional_format_databar_set_negative_fill_color.rs` - Example of
  adding a data bar type conditional formatting to a worksheet with user
  defined negative fill color.

* `doc_conditional_format_databar_set_solid_fill.rs` - Example of adding a
  data bar type conditional formatting to a worksheet with a solid
  (non-gradient) style bar.

* `doc_conditional_format_databar.rs` - Example of adding data bar type
  conditional formatting to a worksheet.

* `doc_conditional_format_date.rs` - Example of adding a Dates Occurring
  type conditional formatting to a worksheet. Note, the rules in this
  example such as "Last month", "This month" and "Next month" are applied
  to the sample dates which by default are for November 2023. Changes the
  dates to some range closer to the time you run the example.

* `doc_conditional_format_duplicate.rs` - Example of how to add a
  duplicate/unique conditional formatting to a worksheet. Duplicate values
  are in light red. Unique values are in light green. Note, that we invert
  the Duplicate rule to get Unique values.

* `doc_conditional_format_error.rs` - Example of how to add a
  error/non-error conditional formatting to a worksheet. Error values are
  in light red. Non-error values are in light green. Note, that we invert
  the Error rule to get Non-error values.

* `doc_conditional_format_formula.rs` - Example of adding a Formula type
  conditional formatting to a worksheet. Cells with odd numbered values are
  in light red while even numbered values are in light green.

* `doc_conditional_format_icon_default_icons.rs` - The following example
  shows Excels default icon settings expressed as `rust_xlsxwriter` rules.

* `doc_conditional_format_icon_reverse_icons.rs` - Example of adding icon
  style conditional formatting to a worksheet. In the second example the
  order of the icons is reversed.

* `doc_conditional_format_icon_set_custom.rs` - Example of adding icon
  style conditional formatting to a worksheet. In the second example the
  default icons are changed.

* `doc_conditional_format_icon_set_icons.rs` - Example of adding icon style
  conditional formatting to a worksheet. In the second example the default
  rules are changed.

* `doc_conditional_format_icon_show_icons_only.rs` - Example of adding icon
  style conditional formatting to a worksheet. In the second example the
  icons are shown without the cell data.

* `doc_conditional_format_icon.rs` - Example of adding icon style
  conditional formatting to a worksheet.

* `doc_conditional_format_multi_range.rs` - Example of adding a cell type
  conditional formatting to a worksheet over a non-contiguous range. Cells
  with values >= 50 are in light red. Values < 50 are in light green. Note
  that the cells outside the selected ranges do not have any conditional
  formatting.

* `doc_conditional_format_text.rs` - Example of adding a text type
  conditional formatting to a worksheet.

* `doc_conditional_format_top.rs` - Example of how to add Top and Bottom
  conditional formatting to a worksheet. Top 10 values are in light red.
  Bottom 10 values are in light green.

* `doc_data_validation_allow_custom.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts input to
  text/strings that are uppercase.

* `doc_data_validation_allow_date.rs` - Example of adding a data validation
  to a worksheet cell. This validation restricts input to date values in a
  fixed range.

* `doc_data_validation_allow_decimal_number_formula.rs` - Example of adding
  a data validation to a worksheet cell. This validation restricts input to
  floating point values based on a value from another cell.

* `doc_data_validation_allow_decimal_number.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts input to
  floating point values in a fixed range.

* `doc_data_validation_allow_list_formula.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts users to a
  selection of values from a dropdown list. The list data is provided from
  a cell range.

* `doc_data_validation_allow_list_strings.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts users to a
  selection of values from a dropdown list.

* `doc_data_validation_allow_list_strings2.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts users to a
  selection of values from a dropdown list. This example shows how to
  pre-populate a default choice.

* `doc_data_validation_allow_text_length.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts input to
  strings whose length is in a fixed range.

* `doc_data_validation_allow_time.rs` - Example of adding a data validation
  to a worksheet cell. This validation restricts input to time values in a
  fixed range.

* `doc_data_validation_allow_whole_number_formula.rs` - Example of adding a
  data validation to a worksheet cell. This validation restricts input to
  integer values based on a value from another cell.

* `doc_data_validation_allow_whole_number_formula2.rs` - Example of adding
  a data validation to a worksheet cell. This validation restricts input to
  integer values based on a value from another cell.

* `doc_data_validation_allow_whole_number.rs` - Example of adding a data
  validation to a worksheet cell. This validation restricts input to
  integer values in a fixed range.

* `doc_data_validation_intro1.rs` - Example of adding a data validation to
  a worksheet cell. This validation uses an input message to explain to the
  user what type of input is required.

* `doc_data_validation_set_error_message.rs` - Example of adding a data
  validation to a worksheet cell. This validation shows a custom error
  message.

* `doc_data_validation_set_error_title.rs` - Example of adding a data
  validation to a worksheet cell. This validation shows a custom error
  title.

* `doc_data_validation_set_input_message.rs` - Example of adding a data
  validation to a worksheet cell. This validation uses an input message to
  explain to the user what type of input is required.

* `doc_datetime_and_hms_milli.rs` - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* `doc_datetime_and_hms.rs` - Demonstrates writing formatted datetimes in
  an Excel worksheet.

* `doc_datetime_from_hms_milli.rs` - Demonstrates writing formatted times
  in an Excel worksheet.

* `doc_datetime_from_hms.rs` - Demonstrates writing formatted times in an
  Excel worksheet.

* `doc_datetime_from_serial_datetime.rs` - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* `doc_datetime_from_timestamp.rs` - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* `doc_datetime_from_ymd.rs` - Demonstrates writing formatted dates in an
  Excel worksheet.

* `doc_datetime_intro.rs` - Demonstrates writing formatted datetimes in an
  Excel worksheet.

* `doc_datetime_parse_from_str.rs` - Demonstrates writing formatted
  datetimes parsed from strings.

* `doc_datetime_to_excel.rs` - Demonstrates the ExcelDateTime `to_excel()`
  method.

* `doc_enum_xlsxcolor.rs` - Demonstrates using different Color enum values
  to set the color of some text in a worksheet.

* `doc_format_clone.rs` - Demonstrates cloning a format and setting the
  properties.

* `doc_format_create.rs` - Demonstrates create a new format and setting the
  properties.

* `doc_format_currency1.rs` - Demonstrates setting a currency format for a
  worksheet cell. This example doesn't actually set a currency format, for
  that see the followup example in doc_format_currency2.rs.

* `doc_format_currency2.rs` - Demonstrates setting a currency format for a
  worksheet cell.

* `doc_format_default.rs` - Demonstrates creating a default format.

* `doc_format_intro.rs` - Demonstrates some of the available formatting
  properties.

* `doc_format_locale.rs` - Demonstrates setting a number format that
  appears differently in different locales.

* `doc_format_merge1.rs` - Demonstrates creating a format that is a
  combination of two formats.

* `doc_format_merge2.rs` - Demonstrates creating a format that is a
  combination of two formats. This example demonstrates that properties in
  the primary format take precedence.

* `doc_format_merge3.rs` - This example demonstrates how cells without
  explicit formats inherit the formats from the row and column that they
  are in. Note the output: - Cell C1 has a green font color. - Cell A3 has
  a bold format. - Cell C3 has both a bold format and a green font color.

* `doc_format_new.rs` - Demonstrates creating a new format.

* `doc_format_set_align.rs` - Demonstrates setting various cell alignment
  properties.

* `doc_format_set_background_color.rs` - Demonstrates setting the cell
  background color, with a default solid pattern.

* `doc_format_set_bold.rs` - Demonstrates setting the bold property for a
  format.

* `doc_format_set_border_color.rs` - Demonstrates setting a cell border and
  color.

* `doc_format_set_border_diagonal.rs` - Demonstrates setting cell diagonal
  borders.

* `doc_format_set_border.rs` - Demonstrates setting a cell border.

* `doc_format_set_font_color.rs` - Demonstrates setting the italic property
  for a format.

* `doc_format_set_font_name.rs` - Demonstrates setting the font name/type
  for a format.

* `doc_format_set_font_size.rs` - Demonstrates setting the font size for a
  format.

* `doc_format_set_font_strikethrough.rs` - Demonstrates setting the text
  strikethrough property for a format.

* `doc_format_set_foreground_color.rs` - Demonstrates setting the
  foreground/pattern color.

* `doc_format_set_indent.rs` - Demonstrates setting the indentation level
  for cell text.

* `doc_format_set_italic.rs` - Demonstrates setting the italic property for
  a format.

* `doc_format_set_num_format_index.rs` - Demonstrates setting one of the
  inbuilt format indices for a format.

* `doc_format_set_num_format.rs` - Demonstrates setting different types of
  Excel number formatting.

* `doc_format_set_pattern.rs` - Demonstrates setting the cell pattern (with
  colors).

* `doc_format_set_quote_prefix.rs` - Demonstrates setting the quote prefix
  property for a format.

* `doc_format_set_reading_direction.rs` - Demonstrates setting the text
  reading direction. This is useful when creating Arabic, Hebrew or other
  near or far eastern worksheets.

* `doc_format_set_rotation.rs` - Demonstrates setting text rotation for a
  cell.

* `doc_format_set_shrink.rs` - Demonstrates setting the text shrink format.

* `doc_format_set_text_wrap.rs` - Demonstrates setting an implicit (without
  newline) text wrap and a user defined text wrap (with newlines).

* `doc_format_set_underline.rs` - Demonstrates setting underline properties
  for a format.

* `doc_image_dimensions.rs` - This example shows how to get some of the
  properties of an Image that will be used in an Excel worksheet.

* `doc_image_new_from_buffer.rs` - This example shows how to create an
  image object from a u8 buffer.

* `doc_image_set_alt_text.rs` - This example shows how to create an image
  object and set the alternative text to help accessibility.

* `doc_image_set_decorative.rs` - This example shows how to create an image
  object and set the decorative property to indicate the it doesn't contain
  useful visual information. This is used to improve the accessibility of
  visual elements.

* `doc_image_set_object_movement.rs` - This example shows how to create an
  image object and set the option to control how it behaves when the cells
  underneath it are changed.

* `doc_image_set_scale_to_size.rs` - An example of scaling images to a
  fixed width and height. See also the
  `worksheet.insert_image_fit_to_cell()` method.

* `doc_image_set_scale_width.rs` - This example shows how to create an
  image object and use it to insert the image into a worksheet. The image
  in this case is scaled.

* `doc_image_set_width.rs` - This example shows how to create an image
  object and use it to insert the image into a worksheet. The image in this
  case is scaled by setting the height and width.

* `doc_image.rs` - This example shows how to create an image object and use
  it to insert the image into a worksheet.

* `doc_into_chart_format.rs` - An example of passing chart formatting
  parameters via the [`IntoChartFormat`] trait.

* `doc_into_color.rs` - An example of the different types of color syntax
  that is supported by the [`Into`] [`Color`] trait.

* `doc_into_shape_format.rs` - An example of passing shape formatting
  parameters via the [`IntoShapeFormat`] trait.

* `doc_macros_add.rs` - Demonstrates a simple example of adding a vba
  project to an xlsm file.

* `doc_macros_calc.rs` - Demonstrates a simple example of adding a vba
  project to an xlsm file.

* `doc_macros_name.rs` - Demonstrates a simple example of adding a vba
  project to an xlsm file.

* `doc_macros_save.rs` - Demonstrates a simple example of adding a vba
  project to an xlsm file.

* `doc_macros_signed.rs` - Demonstrates a simple example of adding a vba
  project to an xlsm file.

* `doc_note_add_author_prefix.rs` - Demonstrates adding a note to a
  worksheet cell. This example turns off the author name in the note.

* `doc_note_new.rs` - Demonstrates adding a note to a worksheet cell.

* `doc_note_reset_text.rs` - Demonstrates adding a note to a worksheet
  cell. This example reuses the Note object and reset the test.

* `doc_note_set_author.rs` - Demonstrates adding a note to a worksheet
  cell. This example also sets the author name.

* `doc_note_set_background_color.rs` - Demonstrates adding a note to a
  worksheet cell. This example also sets the background color.

* `doc_note_set_visible.rs` - Demonstrates adding a note to a worksheet
  cell. This example makes the note visible by default.

* `doc_note_set_width.rs` - Demonstrates adding a note to a worksheet cell.
  This example also changes the note dimensions.

* `doc_properties_checksum1.rs` - Create a simple workbook to demonstrate
  the changing checksum due to the changing creation date.

* `doc_properties_checksum2_chrono.rs` - Create a simple workbook to
  demonstrate a constant checksum due to the a constant creation date.

* `doc_properties_checksum2.rs` - Create a simple workbook to demonstrate a
  constant checksum due to the a constant creation date.

* `doc_properties_custom.rs` - An example of setting custom/user defined
  workbook document properties.

* `doc_shape_font_set_bold.rs` - This example demonstrates adding a Textbox
  shape and setting some of the font properties.

* `doc_shape_font_set_color.rs` - This example demonstrates adding a
  Textbox shape and setting some of the font properties.

* `doc_shape_font_set_italic.rs` - This example demonstrates adding a
  Textbox shape and setting some of the font properties.

* `doc_shape_font_set_name.rs` - This example demonstrates adding a Textbox
  shape and setting some of the font properties.

* `doc_shape_font_set_size.rs` - This example demonstrates adding a Textbox
  shape and setting some of the font properties.

* `doc_shape_format_set_gradient_fill.rs` - This example demonstrates
  adding a Textbox shape and setting some of its properties.

* `doc_shape_format_set_line.rs` - This example demonstrates adding a
  Textbox shape and setting some of the line properties.

* `doc_shape_format_set_no_fill.rs` - This example demonstrates adding a
  Textbox shape and turning off its border.

* `doc_shape_format_set_no_line.rs` - This example demonstrates adding a
  Textbox shape and turning off its border.

* `doc_shape_format.rs` - This example demonstrates adding a Textbox shape
  and setting some of its properties.

* `doc_shape_gradient_fill_set_gradient_stops.rs` - This example
  demonstrates adding a Textbox shape and setting some of the gradient fill
  properties.

* `doc_shape_gradient_fill_set_type.rs` - This example demonstrates adding
  a Textbox shape and setting some of the gradient fill properties.

* `doc_shape_gradient_fill.rs` - This example demonstrates adding a Textbox
  shape and setting some of the gradient fill properties.

* `doc_shape_line_set_color.rs` - This example demonstrates adding a
  Textbox shape and setting some of the line properties.

* `doc_shape_line_set_dash_type.rs` - This example demonstrates adding a
  Textbox shape and setting some of the line properties.

* `doc_shape_line_set_hidden.rs` - This example demonstrates adding a
  Textbox shape and setting some of the line properties.

* `doc_shape_line_set_transparency.rs` - This example demonstrates adding a
  Textbox shape and setting some of the line properties.

* `doc_shape_line_set_width.rs` - This example demonstrates adding a
  Textbox shape and setting some of the line properties.

* `doc_shape_line.rs` - This example demonstrates adding a Textbox shape
  and setting some of the line properties.

* `doc_shape_pattern_fill.rs` - This example demonstrates adding a Textbox
  shape and setting some of the pattern fill properties.

* `doc_shape_set_font.rs` - This example demonstrates adding a Textbox
  shape and setting some of the font properties.

* `doc_shape_set_text_link.rs` - This example demonstrates adding a Textbox
  shape with text from a cell to a worksheet.

* `doc_shape_set_text.rs` - This example demonstrates adding a Textbox
  shape with text to a worksheet.

* `doc_shape_set_width.rs` - This example demonstrates adding a resized
  Textbox shape to a worksheet.

* `doc_shape_solid_fill_set_color.rs` - This example demonstrates adding a
  Textbox shape and setting some of the solid fill properties.

* `doc_shape_solid_fill_set_transparency.rs` - This example demonstrates
  adding a Textbox shape and setting some of the solid fill properties.

* `doc_shape_text_options_set_direction.rs` - This example demonstrates
  adding a Textbox shape and setting some of the text option properties.

* `doc_shape_text_options_set_horizontal_alignment.rs` - This example
  demonstrates adding a Textbox shape and setting some of the text option
  properties. This highlights the difference between horizontal and
  vertical centering.

* `doc_sparkline_set_sparkline_color.rs` - Demonstrates adding a sparkline
  to a worksheet.

* `doc_table_set_alt_text.rs` - Example of adding a worksheet table with
  alt text.

* `doc_table_set_autofilter.rs` - Example of turning off the autofilter in
  a worksheet table.

* `doc_table_set_banded_columns.rs` - Example of turning on the banded
  columns property in a worksheet table. These are normally off by default,

* `doc_table_set_banded_rows.rs` - Example of turning off the banded rows
  property in a worksheet table. These are normally on by default,

* `doc_table_set_columns.rs` - Example of creating a worksheet table.

* `doc_table_set_first_column.rs` - Example of turning on the first column
  highlighting property in a worksheet table. This is normally off by
  default.

* `doc_table_set_header_row.rs` - Example of turning off the default header
  on a worksheet table.

* `doc_table_set_header_row2.rs` - Example of adding a worksheet table with
  a default header.

* `doc_table_set_header_row3.rs` - Example of adding a worksheet table with
  a user defined header captions.

* `doc_table_set_last_column.rs` - Example of turning on the last column
  highlighting property in a worksheet table. This is normally off by
  default.

* `doc_table_set_name.rs` - Example of setting the name of a worksheet
  table.

* `doc_table_set_style.rs` - Example of setting the style of a worksheet
  table.

* `doc_table_set_total_row.rs` - Example of turning on the "totals" row at
  the bottom of a worksheet table. Note, this just turns on the total run
  it doesn't add captions or subtotal functions.

* `doc_table_set_total_row2.rs` - Example of turning on the "totals" row at
  the bottom of a worksheet table with captions and subtotal functions.

* `doc_tablecolumn_set_format.rs` - Example of adding a format to a column
  in a worksheet table.

* `doc_tablecolumn_set_formula.rs` - Example of adding a formula to a
  column in a worksheet table.

* `doc_tablecolumn_set_header_format.rs` - Example of adding a header
  format to a column in a worksheet table.

* `doc_url_intro1.rs` - Demonstrates writing a URL to a worksheet.

* `doc_url_intro2.rs` - Demonstrates writing a URL to a worksheet.

* `doc_url_intro3.rs` - Demonstrates writing a URL to a worksheet.

* `doc_url_set_text.rs` - Demonstrates writing a URL to a worksheet with
  alternative text.

* `doc_utility_check_sheet_name.rs` - Demonstrates testing for a valid
  worksheet name.

* `doc_utility_quote_sheet_name.rs` - Demonstrates quoting worksheet names.

* `doc_workbook_add_worksheet_with_constant_memory.rs` - Demonstrates
  adding worksheets in "standard" and "constant memory" modes.

* `doc_workbook_add_worksheet_with_low_memory.rs` - Demonstrates adding
  worksheets in "standard" and "low memory" modes.

* `doc_workbook_add_worksheet.rs` - Demonstrates creating adding worksheets
  to a workbook.

* `doc_workbook_new.rs` - Demonstrates creating a simple workbook, with one
  unused worksheet.

* `doc_workbook_push_worksheet.rs` - Demonstrates creating a standalone
  worksheet object and then adding it to a workbook.

* `doc_workbook_read_only_recommended.rs` - Demonstrates creating a simple
  workbook which opens with a recommendation that the file should be opened
  in read only mode.

* `doc_workbook_save_to_buffer.rs` - Demonstrates creating a simple
  workbook to a Vec<u8> buffer.

* `doc_workbook_save_to_path.rs` - Demonstrates creating a simple workbook
  using a Rust Path reference.

* `doc_workbook_save_to_writer.rs` - Demonstrates creating a simple
  workbook to some types that implement the `Write` trait like a file and a
  buffer.

* `doc_workbook_save.rs` - Demonstrates creating a simple workbook, with
  one unused worksheet.

* `doc_workbook_set_default_format1.rs` - Demonstrates changing the default
  format for a workbook.

* `doc_workbook_set_default_format2.rs` - Demonstrates changing the default
  format for a workbook.

* `doc_workbook_set_tempdir.rs` - Demonstrates setting a custom directory
  for temporary files when creating a file in "constant memory" mode.

* `doc_workbook_use_custom_theme.rs` - Demonstrates changing the default
  theme for a workbook to a user supplied custom theme.

* `doc_workbook_use_excel_2023_theme.rs` - Demonstrates changing the
  default theme for a workbook. The example uses the Excel 2023
  Office/Aptos theme.

* `doc_workbook_worksheet_from_index.rs` - Demonstrates getting worksheet
  reference by index.

* `doc_workbook_worksheet_from_name.rs` - Demonstrates getting worksheet
  reference by name.

* `doc_workbook_worksheets_mut.rs` - Demonstrates operating on the vector
  of all the worksheets in a workbook.

* `doc_workbook_worksheets.rs` - Demonstrates operating on the vector of
  all the worksheets in a workbook. The non mutable version of this method
  is less useful than `workbook.worksheets_mut()`.

* `doc_working_with_formulas_dynamic_len.rs` - Demonstrates a static
  function which generally returns one value turned into a dynamic function
  which returns a range of values.

* `doc_working_with_formulas_intro.rs` - Demonstrates writing a simple
  formula.

* `doc_working_with_formulas_intro2.rs` - Demonstrates writing a simple
  formula.

* `doc_working_with_formulas_intro3.rs` - Demonstrates writing a simple
  formula.

* `doc_working_with_formulas_static_len.rs` - Demonstrates a static
  function which generally returns one value. Compare this with the dynamic
  function output of doc_working_with_formulas_dynamic_len.rs.

* `doc_working_with_formulas_syntax.rs` - Demonstrates some common formula
  syntax errors.

* `doc_worksheet_add_sparkline_group.rs` - Demonstrates adding a sparkline
  group to a worksheet.

* `doc_worksheet_add_sparkline.rs` - Demonstrates adding a sparkline to a
  worksheet.

* `doc_worksheet_autofilter.rs` - Demonstrates setting a simple autofilter
  in a worksheet.

* `doc_worksheet_autofit.rs` - Demonstrates auto-fitting the worksheet
  column widths based on the data in the columns.

* `doc_worksheet_clear_cell_format.rs` - Demonstrates clearing the
  formatting from some previously written cells in a worksheet.

* `doc_worksheet_clear_cell.rs` - Demonstrates clearing some previously
  written cell data and formatting from a worksheet.

* `doc_worksheet_constant.rs` - Demonstrates adding worksheets in
  "standard", "low memory" and "constant memory" modes.

* `doc_worksheet_deserialize_headers1.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet.

* `doc_worksheet_filter_column1.rs` - Demonstrates setting an autofilter
  with a list filter condition.

* `doc_worksheet_filter_column2.rs` - Demonstrates setting an autofilter
  with multiple list filter conditions.

* `doc_worksheet_filter_column3.rs` - Demonstrates setting an autofilter
  with a list filter for blank cells.

* `doc_worksheet_filter_column4.rs` - Demonstrates setting an autofilter
  with different list filter conditions in separate columns.

* `doc_worksheet_filter_column5.rs` - Demonstrates setting an autofilter
  for a custom number filter.

* `doc_worksheet_filter_column6.rs` - Demonstrates setting an autofilter
  for two custom number filters to create a "between" condition.

* `doc_worksheet_filter_column7.rs` - Demonstrates setting an autofilter to
  show all the non-blank values in a column. This can be done in 2 ways: by
  adding a filter for each district string/number in the column or since
  that may be difficult to figure out programmatically you can set a custom
  filter. Excel uses both of these methods depending on the data being
  filtered.

* `doc_worksheet_group_columns_collapsed1.rs` - An example of how to group
  worksheet columns into outlines with collapsed/hidden rows.

* `doc_worksheet_group_columns_collapsed2.rs` - An example of how to group
  worksheet columns into outlines with collapsed/hidden rows. This example
  shows hows to add secondary groups within a primary grouping. Excel
  requires at least one column between each outline grouping at the same
  level.

* `doc_worksheet_group_columns1.rs` - An example of how to group worksheet
  columns into outlines.

* `doc_worksheet_group_columns2.rs` - An example of how to group worksheet
  columns into outlines. This example shows hows to add secondary groups
  within a primary grouping. Excel requires at least one column between
  each outline grouping at the same level.

* `doc_worksheet_group_rows_collapsed1.rs` - An example of how to group
  worksheet rows into outlines with collapsed/hidden rows.

* `doc_worksheet_group_rows_collapsed2.rs` - An example of how to group
  worksheet rows into outlines with collapsed/hidden rows. This example
  shows hows to add secondary groups within a primary grouping. Excel
  requires at least one row between each outline grouping at the same
  level.

* `doc_worksheet_group_rows_intro1.rs` - An example of how to group
  worksheet rows into outlines.

* `doc_worksheet_group_rows_intro2.rs` - An example of how to group
  worksheet rows into outlines.

* `doc_worksheet_group_rows1.rs` - An example of how to group worksheet
  rows into outlines.

* `doc_worksheet_group_rows2.rs` - An example of how to group worksheet
  rows into outlines. This example shows hows to add secondary groups
  within a primary grouping. Excel requires at least one row between each
  outline grouping at the same level.

* `doc_worksheet_group_symbols_above.rs` - An example of how to group
  worksheet rows into outlines. This example puts the expand/collapse
  symbol above the range for all row groups in the worksheet.

* `doc_worksheet_group_symbols_to_left.rs` - An example of how to group
  worksheet columns into outlines. This example puts the expand/collapse
  symbol to the left of the range for all row groups in the worksheet.

* `doc_worksheet_hide_unused_rows.rs` - Demonstrates efficiently hiding the
  unused rows in a worksheet.

* `doc_worksheet_ignore_error1.rs` - This example demonstrates an Excel
  cell warning.

* `doc_worksheet_insert_chart_with_offset.rs` - Example of adding a chart
  to a worksheet with a pixel offset within the cell.

* `doc_worksheet_insert_checkbox_with_format.rs` - This example
  demonstrates adding adding a checkbox boolean value to a worksheet along
  with a cell format.

* `doc_worksheet_insert_checkbox1.rs` - This example demonstrates adding
  adding checkbox boolean values to a worksheet.

* `doc_worksheet_insert_checkbox2.rs` - This example demonstrates adding
  adding checkbox boolean values to a worksheet by making use of the Excel
  feature that a checkbox is actually a boolean value with a special
  format.

* `doc_worksheet_insert_image_with_offset.rs` - This example shows how to
  add an image to a worksheet at an offset within the cell.

* `doc_worksheet_insert_shape_with_offset.rs` - This example demonstrates
  adding a Textbox shape to a worksheet cell at an offset.

* `doc_worksheet_insert_shape.rs` - This example demonstrates adding a
  Textbox shape to a worksheet.

* `doc_worksheet_name.rs` - Demonstrates getting a worksheet name.

* `doc_worksheet_new.rs` - Demonstrates creating new worksheet objects and
  then adding them to a workbook.

* `doc_worksheet_protect_with_options.rs` - Demonstrates setting the
  worksheet properties to be protected in a protected worksheet. In this
  case we protect the overall worksheet but allow columns and rows to be
  inserted.

* `doc_worksheet_protect_with_password.rs` - Demonstrates protecting a
  worksheet from editing with a password.

* `doc_worksheet_serialize_datetime1.rs` - Demonstrates serializing
  instances of a Serde derived data structure, including datetimes, to a
  worksheet.

* `doc_worksheet_serialize_datetime2.rs` - Demonstrates serializing
  instances of a Serde derived data structure, including chrono datetimes,
  to a worksheet.

* `doc_worksheet_serialize_datetime3.rs` - Example of a serializable struct
  with a Chrono Naive value with a helper function.

* `doc_worksheet_serialize_datetime4.rs` - Demonstrates serializing
  instances of a Serde derived data structure, including `Option` chrono
  datetimes, to a worksheet.

* `doc_worksheet_serialize_datetime5.rs` - Example of a serializable struct
  with an Option Chrono Naive value with a helper function.

* `doc_worksheet_serialize_dimensions1.rs` - Example of getting the
  dimensions of some serialized data. In this example we use the dimensions
  to set a conditional format range.

* `doc_worksheet_serialize_dimensions2.rs` - Example of getting the
  field/column dimensions of some serialized data. In this example we use
  the dimensions to set a conditional format range.

* `doc_worksheet_serialize_headers_custom.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet with custom
  headers and cell formatting.

* `doc_worksheet_serialize_headers_format1.rs` - Demonstrates formatting
  headers during serialization.

* `doc_worksheet_serialize_headers_format2.rs` - Demonstrates formatting
  headers during serialization.

* `doc_worksheet_serialize_headers_format3.rs` - Demonstrates formatting
  cells during serialization.

* `doc_worksheet_serialize_headers_format4.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet with header
  and value formatting.

* `doc_worksheet_serialize_headers_format5.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet with header
  and column formatting.

* `doc_worksheet_serialize_headers_format6.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet with header
  and value formatting.

* `doc_worksheet_serialize_headers_format7.rs` - Demonstrates turning off
  headers during serialization. The example in columns "D:E" have the
  headers turned off.

* `doc_worksheet_serialize_headers_format8.rs` - Demonstrates different
  methods of handling custom properties. The user can either merge them
  with the default properties or use the custom properties exclusively.

* `doc_worksheet_serialize_headers_hide.rs` - Demonstrates serializing data
  without outputting the headers above the data.

* `doc_worksheet_serialize_headers_rename1.rs` - Demonstrates renaming
  fields during serialization by using Serde field attributes.

* `doc_worksheet_serialize_headers_rename2.rs` - Demonstrates renaming
  fields during serialization by specifying custom headers and renaming
  them there.

* `doc_worksheet_serialize_headers_skip1.rs` - Demonstrates skipping fields
  during serialization by using Serde field attributes. Since the field is
  no longer used we also need to tell rustc not emit a `dead_code` warning.

* `doc_worksheet_serialize_headers_skip2.rs` - Demonstrates skipping fields
  during serialization by omitting them from the serialization headers. To
  do this we need to specify custom headers and set
  `use_custom_headers_only()`.

* `doc_worksheet_serialize_headers_skip3.rs` - Demonstrates skipping fields
  during serialization by explicitly skipping them via custom headers.

* `doc_worksheet_serialize_headers_with_options.rs` - Demonstrates
  serializing instances of a Serde derived data structure to a worksheet.

* `doc_worksheet_serialize_headers_with_options2.rs` - Demonstrates
  serializing instances of a Serde derived data structure to a worksheet.

* `doc_worksheet_serialize_headers1.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet.

* `doc_worksheet_serialize_headers2.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet. This
  demonstrates starting the serialization in a different position

* `doc_worksheet_serialize_headers3.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet using
  different methods (both serialization and deserialization).

* `doc_worksheet_serialize_headers4.rs` - Demonstrates serializing
  instances of a Serde derived data structure to a worksheet.

* `doc_worksheet_serialize_intro.rs` - Demonstrates serializing instances
  of a Serde derived data structure to a worksheet.

* `doc_worksheet_serialize_intro2.rs` - Demonstrates serializing instances
  of a Serde derived data structure to a worksheet. This version uses
  header deserialization.

* `doc_worksheet_serialize_table1.rs` - Demonstrates serializing instances
  of a Serde derived data structure to a worksheet with a default worksheet
  table.

* `doc_worksheet_serialize_table2.rs` - Demonstrates serializing instances
  of a Serde derived data structure to a worksheet with a worksheet table
  and a user defined style.

* `doc_worksheet_serialize_table3.rs` - Demonstrates serializing instances
  of a Serde derived data structure to a worksheet with a user defined
  worksheet table.

* `doc_worksheet_serialize_vectors.rs` - Demonstrates serializing instances
  of a Serde derived data structure with vectors to a worksheet.

* `doc_worksheet_serialize.rs` - Demonstrates serializing instances of a
  Serde derived data structure to a worksheet.

* `doc_worksheet_set_active.rs` - Demonstrates setting a worksheet as the
  visible worksheet when a file is opened.

* `doc_worksheet_set_cell_format.rs` - Demonstrates setting the format of a
  worksheet cell separately from writing the cell data.

* `doc_worksheet_set_column_autofit_width.rs` - Demonstrates "auto"-fitting
  the the width of a column in Excel based on the maximum string width. See
  also the [`Worksheet::autofit()`] command.

* `doc_worksheet_set_column_format.rs` - Demonstrates setting the format
  for a column in Excel.

* `doc_worksheet_set_column_hidden.rs` - Demonstrates hiding a worksheet
  column.

* `doc_worksheet_set_column_range_format.rs` - Demonstrates setting the
  format for all the columns in an Excel worksheet. This effectively, and
  efficiently, sets the format for the entire worksheet.

* `doc_worksheet_set_column_width_pixels.rs` - Demonstrates setting the
  width of columns in Excel in pixels.

* `doc_worksheet_set_column_width.rs` - Demonstrates setting the width of
  columns in Excel.

* `doc_worksheet_set_default_note_author.rs` - Demonstrates adding notes to
  a worksheet and setting the default author name.

* `doc_worksheet_set_default_row_height.rs` - Demonstrates setting the
  default row height for all rows in a worksheet.

* `doc_worksheet_set_formula_result_default.rs` - Demonstrates manually
  setting the default result for all non-calculated formulas in a
  worksheet.

* `doc_worksheet_set_formula_result.rs` - Demonstrates manually setting the
  result of a formula. Note, this is only required for non-Excel
  applications that don't calculate formula results.

* `doc_worksheet_set_freeze_panes_top_cell.rs` - Demonstrates setting the
  worksheet panes and also setting the topmost visible cell in the scrolled
  area.

* `doc_worksheet_set_freeze_panes.rs` - Demonstrates setting the worksheet
  panes.

* `doc_worksheet_set_header_image.rs` - Demonstrates adding a header image
  to a worksheet.

* `doc_worksheet_set_header.rs` - Demonstrates setting the worksheet
  header.

* `doc_worksheet_set_hidden.rs` - Demonstrates hiding a worksheet.

* `doc_worksheet_set_landscape.rs` - Demonstrates setting the worksheet
  page orientation to landscape.

* `doc_worksheet_set_margins.rs` - Demonstrates setting the worksheet
  margins.

* `doc_worksheet_set_name.rs` - Demonstrates setting user defined worksheet
  names and the default values when a name isn't set.

* `doc_worksheet_set_nan_string.rs` - Demonstrates handling NaN and
  Infinity values and also setting custom string representations.

* `doc_worksheet_set_page_breaks.rs` - Demonstrates setting page breaks for
  a worksheet.

* `doc_worksheet_set_page_order.rs` - Demonstrates setting the worksheet
  printed page order.

* `doc_worksheet_set_paper.rs` - Demonstrates setting the worksheet paper
  size/type for the printed output.

* `doc_worksheet_set_print_area.rs` - Demonstrates setting the print area
  for several worksheets.

* `doc_worksheet_set_print_first_page_number.rs` - Demonstrates setting the
  page number on the printed page.

* `doc_worksheet_set_print_fit_to_pages.rs` - Demonstrates setting the
  scale of the worksheet to fit a defined number of pages vertically and
  horizontally. This example shows a common use case which is to fit the
  printed output to 1 page wide but have the height be as long as
  necessary.

* `doc_worksheet_set_print_scale.rs` - Demonstrates setting the scale of
  the worksheet page when printed.

* `doc_worksheet_set_range_format_with_border.rs` - Demonstrates setting
  the format for a range of worksheet cells and also adding a border.

* `doc_worksheet_set_range_format.rs` - Demonstrates setting the format of
  worksheet cells separately from writing the cell data.

* `doc_worksheet_set_range_format2.rs` - Demonstrates setting the format of
  worksheet cells when writing the cell data.

* `doc_worksheet_set_repeat_columns.rs` - Demonstrates setting the columns
  to repeat on each printed page.

* `doc_worksheet_set_repeat_rows.rs` - Demonstrates setting the rows to
  repeat on each printed page.

* `doc_worksheet_set_right_to_left.rs` - Demonstrates changing the default
  worksheet and cell text direction changed from left-to-right to
  right-to-left, as required by some middle eastern versions of Excel.

* `doc_worksheet_set_row_format.rs` - Demonstrates setting the format for a
  row in Excel.

* `doc_worksheet_set_row_height_pixels.rs` - Demonstrates setting the
  height for a row in Excel.

* `doc_worksheet_set_row_height.rs` - Demonstrates setting the height for a
  row in Excel.

* `doc_worksheet_set_row_hidden.rs` - Demonstrates hiding a worksheet row.

* `doc_worksheet_set_screen_gridlines.rs` - Demonstrates turn off the
  worksheet worksheet screen gridlines.

* `doc_worksheet_set_selected.rs` - Demonstrates selecting worksheet in a
  workbook. The active worksheet is selected by default so in this example
  the first two worksheets are selected.

* `doc_worksheet_set_selection.rs` - Demonstrates selecting cells in
  worksheets. The order of selection within the range depends on the order
  of `first` and `last`.

* `doc_worksheet_set_tab_color.rs` - Demonstrates set the tab color of
  worksheets.

* `doc_worksheet_set_top_left_cell.rs` - Demonstrates setting the top and
  leftmost visible cell in the worksheet. Often used in conjunction with
  `set_selection()` to activate the same cell.

* `doc_worksheet_set_zoom.rs` - Demonstrates setting the worksheet zoom
  level.

* `doc_worksheet_show_all_notes.rs` - Demonstrates adding notes to a
  worksheet and setting the worksheet property to make them all visible.

* `doc_worksheet_unprotect_range_with_options.rs` - Demonstrates
  unprotecting ranges in a protected worksheet, with additional options.

* `doc_worksheet_unprotect_range.rs` - Demonstrates unprotecting ranges in
  a protected worksheet.

* `doc_worksheet_write_array_formula_with_format.rs` - Demonstrates writing
  an array formulas with formatting to a worksheet.

* `doc_worksheet_write_array_formula.rs` - Demonstrates writing an array
  formulas to a worksheet.

* `doc_worksheet_write_blank.rs` - Demonstrates writing a blank cell with
  formatting, i.e., a cell that has no data but does have formatting.

* `doc_worksheet_write_boolean_with_format.rs` - Demonstrates writing
  formatted boolean values to a worksheet.

* `doc_worksheet_write_boolean.rs` - Demonstrates writing boolean values to
  a worksheet.

* `doc_worksheet_write_column_matrix.rs` - Demonstrates writing an array of
  column arrays to a worksheet.

* `doc_worksheet_write_column.rs` - Demonstrates writing an array of data
  as a column to a worksheet.

* `doc_worksheet_write_date_chrono.rs` - Demonstrates writing formatted
  dates in an Excel worksheet.

* `doc_worksheet_write_date.rs` - Demonstrates writing formatted dates in
  an Excel worksheet.

* `doc_worksheet_write_datetime_chrono.rs` - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* `doc_worksheet_write_datetime_jiff.rs` - Demonstrates writing formatted
  datetimes in an Excel worksheet.

* `doc_worksheet_write_datetime_with_format.rs` - Demonstrates writing
  formatted datetimes in an Excel worksheet.

* `doc_worksheet_write_datetime.rs` - Demonstrates writing datetimes that
  take an implicit format from the column formatting.

* `doc_worksheet_write_dynamic_array_formula_with_format.rs` - Demonstrates
  a static function which generally returns one value turned into a dynamic
  array function which returns a range of values.

* `doc_worksheet_write_dynamic_array_formula.rs` - Demonstrates a static
  function which generally returns one value turned into a dynamic array
  function which returns a range of values.

* `doc_worksheet_write_formula_with_format.rs` - Demonstrates writing
  formulas with formatting to a worksheet.

* `doc_worksheet_write_formula.rs` - Demonstrates writing formulas with
  formatting to a worksheet.

* `doc_worksheet_write_number_with_format.rs` - Demonstrates setting
  different formatting for numbers in an Excel worksheet.

* `doc_worksheet_write_number.rs` - Demonstrates writing unformatted
  numbers to an Excel worksheet. Any numeric type that will convert
  [`Into`] f64 can be transferred to Excel.

* `doc_worksheet_write_rich_string_with_format.rs` - Demonstrates writing a
  "rich" string with multiple formats, and an additional cell format.

* `doc_worksheet_write_rich_string.rs` - Demonstrates writing a "rich"
  string with multiple formats.

* `doc_worksheet_write_row_matrix.rs` - Demonstrates writing an array of
  row arrays to a worksheet.

* `doc_worksheet_write_row.rs` - Demonstrates writing an array of data as a
  row to a worksheet.

* `doc_worksheet_write_string_with_format.rs` - Demonstrates setting
  different formatting for numbers in an Excel worksheet.

* `doc_worksheet_write_string.rs` - Demonstrates writing some UTF-8 strings
  to a worksheet. The UTF-8 encoding is the only encoding supported by the
  Excel file format.

* `doc_worksheet_write_time_chrono.rs` - Demonstrates writing formatted
  times in an Excel worksheet.

* `doc_worksheet_write_time.rs` - Demonstrates writing formatted times in
  an Excel worksheet.

* `doc_worksheet_write_url_with_format.rs` - Demonstrates writing a URL
  with alternative format.

* `doc_worksheet_write_url_with_text.rs` - Demonstrates writing a URL with
  alternative text.

* `doc_xlsxserialize_column_width.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_field_header_format.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_header_format_reuse.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_header_format.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_hide_headers.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_intro.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_num_format.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_rename.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_skip.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_skip2.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_table_default.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_table_style.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_table.rs` - Example of serializing Serde derived
  structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xlsxserialize_value_format.rs` - Example of serializing Serde
  derived structs to an Excel worksheet using `rust_xlsxwriter` and the
  `XlsxSerialize` trait.

* `doc_xmlwriter_perf_test.rs` - Simple performance test to exercise
  xmlwriter without hitting the worksheet::write_data_table() fast path.

