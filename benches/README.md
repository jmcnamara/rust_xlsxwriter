# Benchmarks

This directory contains Criterion benchmarks for `rust_xlsxwriter`.


## Overall Write Performance Benchmark

The `write_performance` benchmark measures the performance of creating xlsx
files for combinations of strings and numbers. Each batch is tested at 1K, 10K,
100K and 1M cells.

The tested combinations are:

- `mixed` (numbers/strings)
- `numbers` (12345.0).
- `string_static`("Foo")
- `string_dynamic` ("Cell (row, col)")

To compare the default worksheet storage method (the entire data structure is
held in memory) against constant/low memory modes:

```bash
# Default worksheet (baseline).
cargo bench --bench write_performance -- --save-baseline default

# Constant memory mode, compared against the baseline.
cargo bench --bench write_performance --features constant_memory -- --baseline default

# Low memory mode, compared against the baseline.
BENCH_LOW_MEMORY=1 cargo bench --bench write_performance --features constant_memory -- --baseline default
```

## Number Formatting Benchmark

The `float_formatting` benchmark compares the performance of writing numeric
cells to Excel worksheets with the Rust default, `ryu`, and `zmij` libraries.

To compare the different libraries:

```bash
# Default worksheet (baseline).
cargo bench --bench float_formatting -- --save-baseline default

# Then compare against the baseline.
cargo bench --bench float_formatting --features ryu  -- --baseline default
cargo bench --bench float_formatting --features zmij -- --baseline default
```
