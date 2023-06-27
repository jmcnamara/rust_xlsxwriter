/*!

A getting started tutorial for `rust_xlsxwriter`.

# Introduction

The `rust_xlsxwriter` library is a Rust library for writing Excel files in
the xlsx format.

<img src="https://rustxlsxwriter.github.io/images/demo.png">

It can be used to write text, numbers, dates and formulas to multiple worksheets
in a new Excel 2007+ xlsx file. It has a focus on performance and on fidelity
with the file format created by Excel.

This document is a tutorial on getting started with `rust_xlsxwriter` and for
using it to write Excel xlsx files.

 # Installation and creating a sample application

In order to use the `rust_xlsxwriter` in a application or in another library you
will need add it as a dependency to the `Cargo.toml` file of your project.

To demonstrate the steps required we will start with a small sample application.

## Create a sample project

Create a new rust command-line application as follows:

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

## Add/install `rust_xlsxwriter`

Add the `rust_xlsxwriter` dependency to the project `Cargo.toml` file:

```bash
$ cargo add rust_xlsxwriter
```

## Modify `main.rs`

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

In order to look at some of the basic but more useful features of the
`rust_xlsxwriter` library will will create an application to summarize some
monthly expenses into a spreadsheet.


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
    for expense in expenses {
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


```ignore
let mut workbook = Workbook::new();
```

**Note**, `rust_xlsxwriter` can only create new files. It cannot read or modify
existing files.

The workbook object is then used to add a new worksheet via the
[`workbook.add_worksheet()`](crate::Workbook::add_worksheet) method:




```ignore
let worksheet = workbook.add_worksheet();
```
The worksheet will have a standard Excel name, in this case "Sheet1". You can
specify the worksheet name using the
[`worksheet.set_name()`](crate::Worksheet::set_name) method.


We then iterate over the data and use the
[`worksheet.write()`](crate::Worksheet::write) method which converts common Rust
types to the equivalent Excel types and writes them to the specified `row, col`
location in the worksheet:

```ignore
    for expense in expenses {
        worksheet.write(row, 0, expense.0)?;
        worksheet.write(row, 1, expense.1)?;
        row += 1;
    }
```

Throughout `rust_xlsxwriter` rows and columns are zero indexed. So, for example,
the first cell in a worksheet, `A1`, is `(0, 0)`.

We then add a [`Formula`](crate::Formula) to calculate the total of the items in
the second column:

```ignore
    worksheet.write(row, 1, Formula::new("=SUM(B1:B4)"))?;
```

Finally, we save and close the Excel file via the
[`workbook.save()`](crate::Workbook::save) method which take a [`std::path`]
`Path`, `PathBuf` or filename string as an argument:

```ignore
    workbook.save("tutorial1.xlsx")?;
```

This will generate the spreadsheet shown in the image above.

It is also possible to save to a byte vector using
[`workbook.save_to_buffer()`](crate::Workbook::save_to_buffer).


## Tutorial Part 2: Adding some formatting

The previous example converted the required data into an Excel file but it
looked a little bare. In order to make the information clearer we can add some
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
    for expense in expenses {
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

`Format` objects represent all of the formatting properties that can be applied
to a cell in Excel such as fonts, number formatting, colors and borders. This is
explained in more detail in the [`Format`](crate::Format) struct documentation.

For now we will avoid getting into the details of Format and just use a limited
amount of the its functionality to add some simple formatting:

```ignore
    // Add a bold format to use to highlight cells.
    let bold = Format::new().set_bold();

    // Add a number format for cells with money values.
    let money_format = Format::new().set_num_format("$#,##0");
```

We can use these formats with the
[`worksheet.write_with_format()`](crate::Worksheet::write_with_format) method
which write data and formatting together.

```ignore
    worksheet.write_with_format(0, 0, "Item", &bold)?;

    worksheet.write_with_format(row, 1, expense.1, &money_format)?;

    worksheet.write_with_format(row, 1, Formula::new("=SUM(B2:B5)"), &money_format)?;
```

## Tutorial Part 3: Adding dates and more formatting

Let's extend the application a little bit more to add some dates to the data:

```rust
    let expenses = vec![
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
    for expense in expenses {
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

Dates and times in Excel are floating point numbers that have a number format
applied to display them in the correct format. In order to handle dates and
times with `rust_xlsxwriter` we create them using a
[`ExcelDateTime`](crate::ExcelDateTime) instance and format them with an Excel
number format.

Alternatively, if you have enable the `chrono` feature in `rust_xlsxwriter`  you
can use [`chrono::NaiveDateTime`], [`chrono::NaiveDate`] or
[`chrono::NaiveTime`] instances instead.

[`chrono::NaiveDate`]:
    https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
[`chrono::NaiveTime`]:
    https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
[`chrono::NaiveDateTime`]:
    https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html

In the example above we create the `ExcelDateTime` instance from the date
strings in our input data and then format it, in Excel, with a number format.

```ignore
        let date = ExcelDateTime::parse_from_str(expense.2)?;
        worksheet.write_date(row, 2, &date, &date_format)?;
```

The final addition to our program is the make the "Date" column wider for
clarity using the
[`worksheet.set_column_width()`](crate::Worksheet.set_column_width) method.


```ignore
    worksheet.set_column_width(2, 15)?;
```

 */
