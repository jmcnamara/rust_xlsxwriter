/*!

A getting started tutorial for `rust_xlsxwriter`.

- [Introduction](#introduction)
- [Getting started](#getting-started)
  - [Create a sample project](#create-a-sample-project)
  - [Add `rust_xlsxwriter` to Cargo.toml](#add-rust_xlsxwriter-to-cargotoml)
  - [Modify main.rs](#modify-mainrs)
  - [Run the application](#run-the-application)
- [Tutorial](#tutorial)
  - [Reading ahead](#reading-ahead)
  - [Tutorial Part 1: Adding data to a worksheet](#tutorial-part-1-adding-data-to-a-worksheet)
  - [Tutorial Part 2: Adding some formatting](#tutorial-part-2-adding-some-formatting)
  - [Tutorial Part 3: Adding dates and more formatting](#tutorial-part-3-adding-dates-and-more-formatting)
  - [Tutorial Part 4: Adding a chart](#tutorial-part-4-adding-a-chart)
  - [Tutorial Part 5: Making the code more programmatic](#tutorial-part-5-making-the-code-more-programmatic)
- [Next steps](#next-steps)


# Introduction

The `rust_xlsxwriter` library is a Rust library for writing Excel files in
the xlsx format.

<img src="https://rustxlsxwriter.github.io/images/demo.png">

It can be used to write text, numbers, dates and formulas to multiple worksheets
in a new Excel 2007+ xlsx file. It has a focus on performance and on fidelity
with the file format created by Excel.

This document is a tutorial on getting started with `rust_xlsxwriter` and for
using it to write Excel xlsx files.

# Getting started

To use the `rust_xlsxwriter` in an application or in another library you
will need add it as a dependency to the `Cargo.toml` file of your project.

To demonstrate the steps required we will start with a small sample application.

## Create a sample project

Create a new Rust command-line application as follows:

```bash
$ cargo new hello-xlsx
```

Change to the new `hello-xlsx` directory:

```bash
$ cd hello-xlsx
```

The directory structure will look like the following:

```bash
hello-xlsx/
├── Cargo.toml
└── src
    └── main.rs
```

## Add `rust_xlsxwriter` to Cargo.toml

Add the `rust_xlsxwriter` dependency to the project `Cargo.toml` file:

```bash
$ cargo add rust_xlsxwriter
```

## Modify main.rs

Modify the `src/main.rs` file so it looks like this:

```rust
// This code is available in examples/app_hello_world.rs
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a string to cell (0, 0) = A1.
    worksheet.write(0, 0, "Hello")?;

    // Write a number to cell (1, 0) = A2.
    worksheet.write(1, 0, 12345)?;

    // Save the file to disk.
    workbook.save("hello.xlsx")?;

    Ok(())
}
```

## Run the application

Run the application as follows:

```bash
$ cargo run
```

This will create an output file called `hello.xlsx` which should look something
like this:

<img src="https://rustxlsxwriter.github.io/images/hello.png">

Once you have the "hello world" application working you can move on to a
slightly more realistic tutorial example.

# Tutorial

To look at some of the basic but more useful features of the
`rust_xlsxwriter` library we will create an application to summarize some
monthly expenses into a spreadsheet.

## Reading ahead

The tutorial presents a simple direct approach so as not to confuse the reader
with information that isn't required for an initial understanding. If there is
more advanced information that might be interesting at a later stage it will be
highlighted in a "Reading ahead" section like this:

> **Reading ahead**:
>
> Some more advanced information.

## Tutorial Part 1: Adding data to a worksheet

To add some sample expense data to a worksheet we could start with a simple
program like the following:

```rust
// This code is available in examples/app_tutorial1.rs
use rust_xlsxwriter::{Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![
        ("Rent", 2000),
        ("Gas", 200),
        ("Food", 500),
        ("Gym", 100),
    ];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Iterate over the data and write it out row by row.
    let mut row = 0;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write(row, 1, expense.1)?;
        row += 1;
    }

    // Write a total using a formula.
    worksheet.write(row, 0, "Total")?;
    worksheet.write(row, 1, Formula::new("=SUM(B1:B4)"))?;

    // Save the file to disk.
    workbook.save("tutorial1.xlsx")?;

    Ok(())
}
```

If we run this program we should get a spreadsheet that looks like this:

<img src="https://rustxlsxwriter.github.io/images/tutorial1.png">

This is a simple program but it demonstrates some of the steps that would apply
to any `rust_xlsxwriter` program.


The first step is to create a new workbook object using the
[`Workbook`](crate::Workbook) constructor
[`Workbook::new()`](crate::Workbook::new):


```text
let mut workbook = Workbook::new();
```

**Note**, `rust_xlsxwriter` can only create new files. It cannot read or modify
existing files.

The workbook object is then used to add a new worksheet via the
[`Workbook::add_worksheet()`](crate::Workbook::add_worksheet) method:

```text
let worksheet = workbook.add_worksheet();
```
The worksheet will have a standard Excel name, in this case "Sheet1". You can
specify the worksheet name using the
[`Worksheet::set_name()`](crate::Worksheet::set_name) method.


We then iterate over the data and use the
[`Worksheet::write()`](crate::Worksheet::write) method which converts common Rust
types to the equivalent Excel types and writes them to the specified `row, col`
location in the worksheet:

```text
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write(row, 1, expense.1)?;
        row += 1;
    }
```

> **Reading ahead**:
>
> There are other type specific write methods such as
> [`Worksheet::write_string()`](crate::Worksheet::write_string) and
> [`Worksheet::write_number()`](crate::Worksheet::write_number). However, these
> aren't generally required and thanks to Rust's monomorphization the
> performance of the generic `write()` method is just as fast.
>
> There are also worksheet methods for writing arrays of data or arrays of
> arrays of data that can be useful in cases where the data to be added is in
> a vector format:
>
> - [`Worksheet::write_row()`](crate::Worksheet::write_row)
> - [`Worksheet::write_column()`](crate::Worksheet::write_column)
> - [`Worksheet::write_row_matrix()`](crate::Worksheet::write_row_matrix)
> - [`Worksheet::write_column_matrix()`](crate::Worksheet::write_column_matrix)
> - [`Worksheet::write_row_with_format()`](crate::Worksheet::write_row_with_format)
> - [`Worksheet::write_column_with_format()`](crate::Worksheet::write_column_with_format)

Throughout `rust_xlsxwriter` rows and columns are zero indexed. So the first
cell in a worksheet `(0, 0)` is equivalent to the Excel notation of `A1`.

To calculate the total of the items in the second column we add a
[`Formula`](crate::Formula):

```text
    worksheet.write(row, 1, Formula::new("=SUM(B1:B4)"))?;
```

Finally, we save and close the Excel file via the
[`Workbook::save()`](crate::Workbook::save) method which will generate the
spreadsheet shown in the image above.:

```text
    workbook.save("tutorial1.xlsx")?;
```

> **Reading ahead**:
>
> The [`Workbook::save()`](crate::Workbook::save) method takes a [`std::path`]
> argument which can be a `Path`, `PathBuf` or a filename string. It is also
> possible to save to a byte vector using
> [`Workbook::save_to_buffer()`](crate::Workbook::save_to_buffer).


## Tutorial Part 2: Adding some formatting

The previous example converted the required data into an Excel file but it
looked a little bare. To make the information clearer we can add some
simple formatting, like this:

<img src="https://rustxlsxwriter.github.io/images/tutorial2.png">


The differences here are that we have added "Item" and "Cost" column headers in
a bold font, we have formatted the currency in the second column and we have
made the "Total" string bold.

To do this programmatically we can extend our code as follows:


```rust
// This code is available in examples/app_tutorial2.rs
use rust_xlsxwriter::{Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![
        ("Rent", 2000),
        ("Gas", 200),
        ("Food", 500),
        ("Gym", 100),
    ];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some column headers.
    worksheet.write_with_format(0, 0, "Item", &bold)?;
    worksheet.write_with_format(0, 1, "Cost", &bold)?;

    // Iterate over the data and write it out row by row.
    let mut row = 1;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write_with_format(row, 1, expense.1, &money_format)?;
        row += 1;
    }

    // Write a total using a formula.
    worksheet.write_with_format(row, 0, "Total", &bold)?;
    worksheet.write_with_format(row, 1, Formula::new("=SUM(B2:B5)"), &money_format)?;

    // Save the file to disk.
    workbook.save("tutorial2.xlsx")?;

    Ok(())
}
```

The main difference between this and the previous program is that we have added
two [`Format`](crate::Format) objects that we can use to format cells in the
spreadsheet.

`Format` objects represent all the formatting properties that can be applied
to a cell in Excel such as fonts, number formatting, colors and borders. This is
explained in more detail in the [`Format`](crate::Format) struct documentation.

For now we will avoid getting into the details of `Format` and just use a
limited amount of the its functionality to add some simple formatting:

```text
    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");
```

We can use these formats with the
[`Worksheet::write_with_format()`](crate::Worksheet::write_with_format) method
which writes data and formatting together, like these examples from the code:

```text
    worksheet.write_with_format(0, 0, "Item", &bold)?;

    worksheet.write_with_format(row, 1, expense.1, &money_format)?;

    worksheet.write_with_format(row, 1, Formula::new("=SUM(B2:B5)"), &money_format)?;
```

## Tutorial Part 3: Adding dates and more formatting

Let's extend the application a little bit more to add some dates to the data:

```rust
    let expenses = [
        ("Rent", 2000, "2022-09-01"),
        ("Gas", 200, "2022-09-05"),
        ("Food", 500, "2022-09-21"),
        ("Gym", 100, "2022-09-28"),
    ];
```

The corresponding spreadsheet will look like this:

<img src="https://rustxlsxwriter.github.io/images/tutorial3.png">

The differences here are that we have added a "Date" column with formatting and
made that column a little wider to accommodate the dates.

To do this we can extend our program as follows:

```rust
// This code is available in examples/app_tutorial3.rs
use rust_xlsxwriter::{ExcelDateTime, Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![
        ("Rent", 2000, "2022-09-01"),
        ("Gas", 200, "2022-09-05"),
        ("Food", 500, "2022-09-21"),
        ("Gym", 100, "2022-09-28"),
    ];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");

    // Add a number format for cells with dates.
    let date_format = Format::new().set_num_format("d mmm yyyy");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some column headers.
    worksheet.write_with_format(0, 0, "Item", &bold)?;
    worksheet.write_with_format(0, 1, "Cost", &bold)?;
    worksheet.write_with_format(0, 2, "Date", &bold)?;

    // Adjust the date column width for clarity.
    worksheet.set_column_width(2, 15)?;

    // Iterate over the data and write it out row by row.
    let mut row = 1;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write_with_format(row, 1, expense.1, &money_format)?;

        let date = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_with_format(row, 2, &date, &date_format)?;

        row += 1;
    }

    // Write a total using a formula.
    worksheet.write_with_format(row, 0, "Total", &bold)?;
    worksheet.write_with_format(row, 1, Formula::new("=SUM(B2:B5)"), &money_format)?;

    // Save the file to disk.
    workbook.save("tutorial3.xlsx")?;

    Ok(())
}
```

Dates and times in Excel are floating point numbers that have a format applied
to display them in the desired way. To handle dates and times with
`rust_xlsxwriter` we create them using a [`ExcelDateTime`](crate::ExcelDateTime)
instance and format them with an Excel number format.

> **Reading ahead**:

> If you enable the `chrono` feature in `rust_xlsxwriter`  you can also use
> [`chrono::NaiveDateTime`], [`chrono::NaiveDate`] or [`chrono::NaiveTime`]
> instances.

[`chrono::NaiveDate`]:
    https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
[`chrono::NaiveTime`]:
    https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
[`chrono::NaiveDateTime`]:
    https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html

In the example above we create the `ExcelDateTime` instance from the date
strings in our input data and then add a number format it so that it appears
correctly in Excel:

```text
        let date = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_with_format(row, 2, &date, &date_format)?;
```

Another addition to our program is the make the "Date" column wider for
clarity using the
[`Worksheet::set_column_width()`](crate::Worksheet.set_column_width) method.


```text
    worksheet.set_column_width(2, 15)?;
```


## Tutorial Part 4: Adding a chart

To extend our example a little further let's add a Pie chart to show the
relative sizes of the outgoing expenses to get a spreadsheet that will look like
this:

<img src="https://rustxlsxwriter.github.io/images/tutorial4.png">

We use the [`Chart`](crate::Chart) struct to represent the chart.

The [`Chart`](crate::Chart) struct has a lot of configuration options and
sub-structs to replicate Excel's chart features but as an initial demonstration
we will just add the data series to which the chart refers. Here is the updated
code with the chart addition at the end.


```rust
// This code is available in examples/app_tutorial4.rs
use rust_xlsxwriter::{Chart, ExcelDateTime, Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![
        ("Rent", 2000, "2022-09-01"),
        ("Gas", 200, "2022-09-05"),
        ("Food", 500, "2022-09-21"),
        ("Gym", 100, "2022-09-28"),
    ];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");

    // Add a number format for cells with dates.
    let date_format = Format::new().set_num_format("d mmm yyyy");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some column headers.
    worksheet.write_with_format(0, 0, "Item", &bold)?;
    worksheet.write_with_format(0, 1, "Cost", &bold)?;
    worksheet.write_with_format(0, 2, "Date", &bold)?;

    // Adjust the date column width for clarity.
    worksheet.set_column_width(2, 15)?;

    // Iterate over the data and write it out row by row.
    let mut row = 1;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write_with_format(row, 1, expense.1, &money_format)?;

        let date = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_with_format(row, 2, &date, &date_format)?;

        row += 1;
    }

    // Write a total using a formula.
    worksheet.write_with_format(row, 0, "Total", &bold)?;
    worksheet.write_with_format(row, 1, Formula::new("=SUM(B2:B5)"), &money_format)?;

    // Add a chart to display the expenses.
    let mut chart = Chart::new_pie();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 4, &chart)?;

    // Save the file to disk.
    workbook.save("tutorial4.xlsx")?;

    Ok(())
}
```

See the documentation for [`Chart`](crate::Chart) for more information.

## Tutorial Part 5: Making the code more programmatic

The previous example worked as expected but it it contained some hard-coded cell
ranges like `set_values("Sheet1!$B$2:$B$5")` and `Formula::new("=SUM(B2:B5)")`.
If our example changed to have a different number of data items then we would
have to manually change the code to adjust for the new ranges.

Fortunately, these hard-coded values are only used for the sake of a tutorial and
`rust_xlsxwriter` provides APIs to handle these more programmatically.

Let's start by looking at the chart ranges:

```text
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");
```

In general `rust_xlsxwriter` always provides numeric APIs for any ranges in
Excel but when it makes ergonomic sense it also provides **secondary** string
based APIs. The previous example uses one of these secondary string based APIs
for demonstration purposes but for real applications you would set the chart
ranges using 5-tuple values like this:

```text
    chart
        .add_series()
        .set_categories(("Sheet1", first_row, item_col, last_row, item_col))
        .set_values(("Sheet1", first_row, cost_col, last_row, cost_col));
```

Where the range values are set or calculated in the code with something like the
following:

```text
    let first_row = 1;
    let last_row = first_row + (expenses.len() as u32) - 1;
    let item_col = 0;
    let cost_col = 1;
```

This allows the range to change dynamically if we add new elements to our `data`
vector and ensures that the worksheet name is quoted properly (when
required).

The other section of the code that had a hard-coded string is the formula
`"=SUM(B2:B5)"`. There isn't a single API change that can be applied to ranges
in formulas but `rust_xlsxwriter` provides several utility functions that can
convert numbers to string ranges. For example the
[`cell_range()`](crate::utility::cell_range) function which takes zero indexed
numbers and converts them to a string range like `B2:B5`:


```text
    let range = cell_range(first_row, cost_col, last_row, cost_col);
    let formula = format!("=SUM({range})");
    worksheet.write_with_format(row, 1, Formula::new(formula), &money_format)?;
```


> **Reading ahead**:
>
> The `cell_range()` function and other similar functions are detailed in the
> [`utility`](crate::utility) documentation.

Adding these improvements our application changes to the following:

```rust
// This code is available in examples/app_tutorial5.rs
use rust_xlsxwriter::{cell_range, Chart, ExcelDateTime, Format, Formula, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Some sample data we want to write to a spreadsheet.
    let expenses = vec![
        ("Rent", 2000, "2022-09-01"),
        ("Gas", 200, "2022-09-05"),
        ("Food", 500, "2022-09-21"),
        ("Gym", 100, "2022-09-28"),
    ];

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");

    // Add a number format for cells with dates.
    let date_format = Format::new().set_num_format("d mmm yyyy");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some column headers.
    worksheet.write_with_format(0, 0, "Item", &bold)?;
    worksheet.write_with_format(0, 1, "Cost", &bold)?;
    worksheet.write_with_format(0, 2, "Date", &bold)?;

    // Adjust the date column width for clarity.
    worksheet.set_column_width(2, 15)?;

    // Iterate over the data and write it out row by row.
    let mut row = 1;
    for expense in &expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write_with_format(row, 1, expense.1, &money_format)?;

        let date = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_with_format(row, 2, &date, &date_format)?;

        row += 1;
    }

    // For clarity, define some variables to use in the formula and chart ranges.
    // Row and column numbers are all zero-indexed.
    let first_row = 1; // Skip the header row.
    let last_row = first_row + (expenses.len() as u32) - 1;
    let item_col = 0;
    let cost_col = 1;

    // Write a total using a formula.
    worksheet.write_with_format(row, 0, "Total", &bold)?;

    let range = cell_range(first_row, cost_col, last_row, cost_col);
    let formula = format!("=SUM({range})");
    worksheet.write_with_format(row, 1, Formula::new(formula), &money_format)?;

    // Add a chart to display the expenses.
    let mut chart = Chart::new_pie();

    // Configure the data series for the chart.
    chart
        .add_series()
        .set_categories(("Sheet1", first_row, item_col, last_row, item_col))
        .set_values(("Sheet1", first_row, cost_col, last_row, cost_col));

    // Add the chart to the worksheet.
    worksheet.insert_chart(1, 4, &chart)?;

    // Save the file to disk.
    workbook.save("tutorial5.xlsx")?;

    Ok(())
}
```

This gives the same output to the previous version, but it is now future proof
for any changes to our input data:

<img src="https://rustxlsxwriter.github.io/images/tutorial5.png">


# Next steps

Once you have completed this simple application you can looks through the
following resources for more information:

- The main API [documentation for the library](../index.html). Click on the
  `rust_xlsxwriter` logo at the top left to get back to the start.

- [Cookbook](crate::cookbook) to see other examples that you can use to get you
  started.

- The [`Workbook`](crate::Workbook) APIs and introduction.
- The [`Worksheet`](crate::Worksheet) APIs and introduction.
- The [`Format`](crate::Format) APIs and introduction.
- The [`Chart`](crate::Chart) APIs and introduction.
- The [`Table`](crate::Table) APIs and introduction.

- [User Guide]: Some longer explanations of parts of `rust_xlsxwriter` API.

- [Release Notes].
- [Roadmap of planned features].

[Release Notes]: crate::changelog
[User Guide]: https://rustxlsxwriter.github.io/index.html
[Roadmap of planned features]:
    https://github.com/jmcnamara/rust_xlsxwriter/issues/1

*/
