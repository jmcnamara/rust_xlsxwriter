/*!

# Performance characteristics of `rust_xlsxwriter`

This section explains some of the performance characteristics of `rust_xlsxwriter`.

Contents:

- [Performance characteristics of `rust_xlsxwriter`](#performance-characteristics-of-rust_xlsxwriter)
  - [Relative performance of `rust_xlsxwriter`](#relative-performance-of-rust_xlsxwriter)
  - [Performance profile](#performance-profile)
  - [Constant memory mode](#constant-memory-mode)
    - [Enabling "constant memory" mode](#enabling-constant-memory-mode)
    - [Restrictions when using "constant memory" mode](#restrictions-when-using-constant-memory-mode)
  - [Zip compression level](#zip-compression-level)
  - [ZMIJ - faster floating-point conversion](#zmij---faster-floating-point-conversion)
  - [Performance testing](#performance-testing)
  - [Programs used to generate the test results](#programs-used-to-generate-the-test-results)
    - [Rust performance test program](#rust-performance-test-program)
    - [C performance test program](#c-performance-test-program)
    - [Python performance test program](#python-performance-test-program)

## Relative performance of `rust_xlsxwriter`

The `rust_xlsxwriter` library has sister libraries written in C
([libxlsxwriter]), Python ([XlsxWriter]), and Perl ([Excel::Writer::XLSX]).

A relative performance comparison between the C, Rust, and Python versions is:

- `rust_xlsxwriter` is 1.6x faster than C/libxlsxwriter.
- `rust_xlsxwriter` is 6.5x faster than Python/XlsxWriter.

The Perl performance is similar to the Python library.

The programs and methodology to generate this data are shown in the section below
on [Performance testing](#performance-testing).

[XlsxWriter]: https://xlsxwriter.readthedocs.io/index.html
[libxlsxwriter]: https://libxlsxwriter.github.io
[Excel::Writer::XLSX]: https://metacpan.org/dist/Excel-Writer-XLSX/view/lib/Excel/Writer/XLSX.pm


## Performance profile

The `rust_xlsxwriter` crate has a linear performance and memory profile relative
to the number of cells written. In general, it will take twice as long and
use twice as much memory to write twice as many cells. Here is a sample speed
performance profile from 100,000-1,000,000 cells:

<img src="https://rustxlsxwriter.github.io/images/performance_speed1.png">

The equivalent memory profile is shown in the next section.

## Constant memory mode

The `rust_xlsxwriter` library maintains an in-memory structure that represents
the cells of a worksheet. This has many usability benefits, such as allowing
the user to apply formatting separately from the data with methods like
[`Worksheet::set_cell_format()`](crate::Worksheet::set_cell_format), or to
perform actions such as [`Worksheet::autofit()`](crate::Worksheet::autofit). The
downside is that memory usage increases (more or less linearly) with the number
of cells that are written.

For most applications you will be able to handle hundreds of millions of
cells before this becomes noticeable at a system level. However, if required, it
is possible to limit the amount of memory used by `rust_xlsxwriter` by using the
`constant_memory` crate-level feature.

The `constant_memory` mode works by flushing the current row of data to disk
when the user writes to a new row of data. This limits the overhead to one row
of data stored in memory. Once this happens, it is no longer possible to write
to a previous row since the data in the Excel file must be in row order. As
such, this imposes the limitation of having to structure your code to write in
row-by-row order. The benefit is that the required memory usage is very low and
effectively constant, regardless of the amount of data written.

| Cells   | Standard Memory | Constant memory |
|---------|-----------------|-----------------|
| 100000  | 18.0            | 0.0215          |
| 200000  | 36.2            | 0.0215          |
| 300000  | 60.1            | 0.0215          |
| 400000  | 72.5            | 0.0215          |
| 500000  | 91.3            | 0.0215          |
| 600000  | 120.5           | 0.0215          |
| 700000  | 132.9           | 0.0215          |
| 800000  | 145.3           | 0.0215          |
| 900000  | 157.7           | 0.0215          |
| 1000000 | 216.8           | 0.0215          |

The `constant_memory` mode also uses an Excel optimization to store string data
"inline" in the cell data rather than in a "shared string table" where only
unique string references are stored (similar to a hash table). This can increase
the final file size if it contains a lot of string data. As a compromise,
`rust_xlsxwriter` also supports a "low memory" mode where only one row of data
is kept in memory and the "shared string table" is used to store unique strings
in memory until the file is written. This keeps memory usage as low as possible
but maintains smaller file sizes and compatibility with standard Excel output.

The table below shows the memory profile when generating a worksheet with 50%
unique string and 50% numerical data in "low memory" mode.

| Cells   | Standard Memory | Low memory |
|---------|-----------------|------------|
| 100000  | 18.0            | 3.0        |
| 200000  | 36.2            | 6.2        |
| 300000  | 60.1            | 11.0       |
| 400000  | 72.5            | 12.5       |
| 500000  | 91.3            | 20.6       |
| 600000  | 120.5           | 22.2       |
| 700000  | 132.9           | 23.8       |
| 800000  | 145.3           | 25.4       |
| 900000  | 157.7           | 27.0       |
| 1000000 | 216.8           | 41.7       |

The memory usage of "low memory" mode will approach the memory usage level of
"constant memory" mode as the percentage of unique string data decreases.

The graphs below show the memory usage and speed for "standard", "low memory"
and "constant memory" modes:

<img src="https://rustxlsxwriter.github.io/images/performance_memory1.png">

<br>
The speed performance is similar in all modes but slightly better (around 5%)
in "constant memory" mode. For worksheets with a large number of unique strings,
"constant memory" mode can be up to 30% faster since it doesn't need to
maintain a shared string table.
<br>
<br>

<img src="https://rustxlsxwriter.github.io/images/performance_speed2.png">


### Enabling "constant memory" mode

To enable "constant memory" mode, you need to add `rust_xlsxwriter` to your
project with the `constant_memory` feature enabled.

```bash
cargo add rust_xlsxwriter -F constant_memory
```

You can then add worksheets for "standard", "low memory"
and "constant memory" modes:


```ignore
# // This code is available in examples/doc_worksheet_constant.rs
#
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet in standard mode.
    let worksheet = workbook.add_worksheet();
    worksheet.write(0, 0, "Standard")?;

    // Add a worksheet in "constant memory" mode.
    let worksheet = workbook.add_worksheet_with_constant_memory()?;
    worksheet.write(0, 0, "Constant memory")?;

    // Add a worksheet in "low memory" mode.
    let worksheet = workbook.add_worksheet_with_low_memory()?;
    worksheet.write(0, 0, "Low memory")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
```

The output file looks like the following:

<img src="https://rustxlsxwriter.github.io/images/worksheet_constant.png">

The ability to add different types/modes of worksheets means that you can mix
smaller random-access worksheets with larger row-by-row worksheets. See the
following for more information:

- [`Worksheet::add_worksheet()`]
- `Worksheet::add_worksheet_with_constant_memory()`
- `Worksheet::add_worksheet_with_low_memory()`

[`Worksheet::add_worksheet()`]: crate::Workbook::add_worksheet

### Restrictions when using "constant memory" mode

There are some limitations and restrictions when using "constant memory" mode.

- Data must be written in row-by-row order, and when you write to row `n`, you can
  no longer write to any row `< n`.
- Constant memory mode uses a [tempfile] filehandle for each worksheet created
  using `Worksheet::add_worksheet_with_constant_memory()` and
  `Worksheet::add_worksheet_with_low_memory()`. This won't save memory if your
  temp directory is also mounted in memory. However, you can set the temp
  directory to a custom location using `Workbook::set_tempdir()`.
- Functions that set formatting separately from data, such as
  [`Worksheet::set_cell_format()`](crate::Worksheet::set_cell_format), will
  only work on the current row.

[tempfile]: https://crates.io/crates/tempfile


## Zip compression level

An xlsx file is a zip container of XML files, and a large part of the time taken
to save a file is spent compressing the XML data and in particular the worksheet
data. By default `rust_xlsxwriter` uses a compression level of 6, which is the
default level for the `zip.rs` crate and which give a good balance between file
size and speed.

If required, the compression level can be changed using
[`Workbook::set_zip_compression_level()`] to trade file size for speed:

[`Workbook::set_zip_compression_level()`]: crate::Workbook::set_zip_compression_level


```rust
# // This code is available in examples/doc_workbook_set_zip_compression_level.rs
#
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // Write some data to the worksheet.
    for row in 0..1000 {
        worksheet.write_number(row, 0, row)?;
        worksheet.write_string(row, 1, "Hello")?;
    }

    // Use a faster compression level than the default level 6.
    workbook.set_zip_compression_level(3);

    workbook.save("workbook.xlsx")?;

    Ok(())
}
```

The level can be in the range 1 (fastest) to 9 (nominally smallest). The table
below shows the file size and save time for each compression level relative to
the default level 6. The worksheet used in this analysis had 4,000 rows x 50
columns of alternating numbers and strings, with 2,000 distinct values of each,
and only the time to save the file is measured.

| Level | File size (bytes) | Size vs 6 | Save time vs 6 |
|-------|-------------------|-----------|----------------|
| 1     | 2,052,818         | 1.51      | 0.45           |
| 2     | 1,502,037         | 1.11      | 0.57           |
| 3     | 1,441,286         | 1.06      | 0.71           |
| 4     | 1,381,297         | 1.02      | 0.86           |
| 5     | 1,364,029         | 1.00      | 0.92           |
| 6     | 1,358,017         | 1.00      | 1.00           |
| 7     | 1,322,184         | 0.97      | 1.53           |
| 8     | 1,314,882         | 0.97      | 6.48           |
| 9     | 1,356,746         | 1.00      | 1.86           |

<br>

If you are looking to increase performance then level 3 is probably the best
compromise. It give ~30% better performance for a file size that is ~5% bigger
but comparable in size to a similar file produced by Excel. Excel 365 also uses
a compression level of 3.


## ZMIJ - faster floating-point conversion

The `rust_xlsxwriter` `zmij` feature flag enables the [zmij] crate, which provides
a "pure Rust implementation of Żmij, an algorithm to quickly convert
floating-point numbers to decimal strings".

This speeds up writing numeric worksheet by about 10%.

<img src="https://rustxlsxwriter.github.io/images/performance_speed4.png">

[zmij]: https://crates.io/crates/zmij/


## Performance testing

The [hyperfine] application was used to run the performance comparison between
the C, Rust, and Python versions discussed above in [Relative performance of
`rust_xlsxwriter`](#relative-performance-of-rust_xlsxwriter).



The main test was as follows:

[hyperfine]: https://lib.rs/crates/hyperfine


```bash
$ hyperfine ./c_perf_test              \
            ./rust_perf_test           \
            "python py_perf_test.py"   \
            --warmup 3 --sort command

Benchmark 1: ./c_perf_test
  Time (mean ± σ):     213.8 ms ±   6.9 ms    [User: 173.8 ms, System: 33.3 ms]
  Range (min … max):   203.5 ms … 223.8 ms    13 runs

Benchmark 2: ./rust_perf_test
  Time (mean ± σ):     135.2 ms ±   5.5 ms    [User: 117.7 ms, System: 15.6 ms]
  Range (min … max):   128.3 ms … 146.1 ms    21 runs

Benchmark 3: python py_perf_test.py
  Time (mean ± σ):     878.2 ms ±  24.8 ms    [User: 822.5 ms, System: 46.1 ms]
  Range (min … max):   837.4 ms … 907.6 ms    10 runs

Relative speed comparison
        1.58 ±  0.08  ./c_perf_test
        1.00          ./rust_perf_test
        6.49 ±  0.32  python py_perf_test.py
```

This shows that `rust_xlsxwriter` is the fastest version and that it is 1.58
times faster than the C version and 6.5 times faster than the Python version.

As with any performance test, there are many factors that may affect the
results. However, these results are indicative of the relative performance.

The programs used to generate these results are shown below.




## Programs used to generate the test results

### Rust performance test program

```rust
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {

    let col_max = 50;
    let row_max = 4_000;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for row in 0..row_max {
        for col in 0..col_max {
            if col % 2 == 1 {
                worksheet.write_string(row, col, "Foo")?;
            } else {
                worksheet.write_number(row, col, 12345.0)?;
            }
        }
    }
    workbook.save("rust_perf_test.xlsx")?;

    Ok(())
}
```

### C performance test program

```C
#include "xlsxwriter.h"

int main() {

    int max_row = 4000;
    int max_col = 50;

    lxw_workbook  *workbook  = workbook_new("c_perf_test.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    for (int row_num = 0; row_num < max_row; row_num++) {
        for (int col_num = 0; col_num < max_col; col_num++) {
            if (col_num % 2)
                worksheet_write_string(worksheet, row_num, col_num, "Foo", NULL);
            else
                worksheet_write_number(worksheet, row_num, col_num, 12345.0, NULL);

        }
    }

    workbook_close(workbook);

    return 0;
}
```

### Python performance test program

```python
import xlsxwriter

row_max = 4000
col_max = 50

workbook = xlsxwriter.Workbook('py_perf_test.xlsx')
worksheet = workbook.add_worksheet()

for row in range(0, row_max):
    for col in range(0, col_max):
        if col % 2:
            worksheet.write_string(row, col, "Foo")
        else:
            worksheet.write_number(row, col, 12345)

workbook.close()
```

*/
