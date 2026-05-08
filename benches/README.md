# Benchmarks

This directory contains Criterion benchmarks for `rust_xlsxwriter`.

## Number Formatting Benchmark

The `number_formatting` benchmark compares the performance of writing numeric
cells to Excel worksheets with different number formatting libraries.

### Running the Benchmarks

To compare the performance of default, `ryu`, and `zmij` features, run the
benchmark multiple times with different features:

#### 1. Default (standard Rust formatting)
```bash
cargo bench --bench number_formatting
```

#### 2. With `ryu` feature
```bash
cargo bench --bench number_formatting --features ryu
```

#### 3. With `zmij` feature
```bash
cargo bench --bench number_formatting --features zmij
```

### Comparing Results

Criterion automatically compares results between runs. After running the
baseline (default), subsequent runs with features will show performance
differences.

To establish a new baseline:
```bash
cargo bench --bench number_formatting -- --save-baseline default
```

Then compare against it:
```bash
cargo bench --bench number_formatting --features ryu -- --baseline default
cargo bench --bench number_formatting --features zmij -- --baseline default
```

### Benchmark Groups

The benchmark includes three groups:

1. **write_numbers**: Writes only numeric values to a worksheet and saves to
   disk:
   - Tests: 5K, 20K, 50K, 200K, and 500K cells

2. **write_numbers_to_memory**: Writes numeric values and saves to memory buffer
   (faster, focuses on formatting):
   - Tests: 5K, 20K, 50K, and 200K cells

3. **write_mixed_data**: Writes alternating numbers and strings:
   - Tests: 5K and 50K cells
