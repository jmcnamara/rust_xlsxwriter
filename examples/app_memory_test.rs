// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! Simple performance test and memory usage program for rust_xlsxwriter.
//!
//! It writes alternate cells of strings and numbers.
//! It defaults to 4,000 rows x 40 columns.
//!
//! The number of rows and the "constant memory" mode can be optionally set.
//!
//! usage: ./target/release/examples/app_perf_test [num_rows]
//! [--constant-memory]
//!

use rust_xlsxwriter::{Workbook, XlsxError};
use std::{env, time::Instant};

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicU64, Ordering};

// The following is used to calculate the memory usage of the program. See:
// https://stackoverflow.com/a/71889391/10238
// https://github.com/discordance/trallocator/blob/master/src/lib.rs

pub struct Trallocator<A: GlobalAlloc>(pub A, AtomicU64);

unsafe impl<A: GlobalAlloc> GlobalAlloc for Trallocator<A> {
    unsafe fn alloc(&self, l: Layout) -> *mut u8 {
        self.1.fetch_add(l.size() as u64, Ordering::SeqCst);
        self.0.alloc(l)
    }
    unsafe fn dealloc(&self, ptr: *mut u8, l: Layout) {
        self.0.dealloc(ptr, l);
        self.1.fetch_sub(l.size() as u64, Ordering::SeqCst);
    }
}

impl<A: GlobalAlloc> Trallocator<A> {
    pub const fn new(a: A) -> Self {
        Trallocator(a, AtomicU64::new(0))
    }

    pub fn reset(&self) {
        self.1.store(0, Ordering::SeqCst);
    }
    pub fn get(&self) -> u64 {
        self.1.load(Ordering::SeqCst)
    }
}

#[global_allocator]
static GLOBAL: Trallocator<System> = Trallocator::new(System);

// The program to test.
fn main() -> Result<(), XlsxError> {
    let args: Vec<String> = env::args().collect();

    // Set some size arguments, optionally from the command line.
    let col_max = 50;
    let row_max = match args.get(1) {
        Some(arg) => arg.parse::<u32>().unwrap_or(4_000),
        None => 4_000,
    };
    let constant_memory = args.get(2).is_some();

    GLOBAL.reset();
    let start_time = Instant::now();

    // Create the workbook and fill in the required cell data.
    let mut workbook = Workbook::new();
    let worksheet = workbook
        .add_worksheet()
        .set_constant_memory_mode(constant_memory)?;

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

    // Calculate and print the metrics.
    let time = (start_time.elapsed().as_millis() as f64) / 1000.0;
    let memory = (GLOBAL.get() as f64) / 1_000_000.0;

    println!("Wrote:  {row_max} rows x {col_max} cols. Constant memory = {constant_memory}.");
    println!("Time:   {time:.3} seconds.");
    println!("Memory: {memory:.3} MB.");

    Ok(())
}
