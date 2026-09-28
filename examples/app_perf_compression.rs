// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! Performance test program for the `rust_xlsxwriter` zip compression levels.
//!
//! It writes a worksheet of alternate string and number cells and then times
//! `save_to_buffer()` at each compression level from 1 to 9. The worksheet
//! data is only created once so only the save/compression time is measured.
//!
//! The levels are interleaved within each round so that any machine drift
//! affects all levels equally, and the first round is a discarded warm-up. The
//! output is CSV with the median, min and 10th/90th percentile save times, and
//! the size and time relative to level 6, which is the default level.
//!
//! The data mode can be "repetitive", which uses the same data as
//! `app_perf_test`, or "varied", which uses a deterministic mix of 2,000
//! distinct strings and numbers for more realistic compression. It defaults to
//! "repetitive", 4,000 rows x 50 columns and 20 rounds.
//!
//! usage: ./target/release/examples/app_perf_compression [repetitive|varied] [num_rows] [rounds]
//!

use rust_xlsxwriter::{Workbook, XlsxError};
use std::{env, time::Instant};

fn main() -> Result<(), XlsxError> {
    let args: Vec<String> = env::args().collect();
    let mode = args.get(1).map_or("repetitive", String::as_str);
    let row_max: u32 = args.get(2).and_then(|a| a.parse().ok()).unwrap_or(4_000);
    let rounds: usize = args.get(3).and_then(|a| a.parse().ok()).unwrap_or(20);
    let col_max = 50;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Simple deterministic LCG so the "varied" data is the same on every run.
    let mut seed: u64 = 12345;
    let mut next = || {
        seed = seed.wrapping_mul(6_364_136_223_846_793_005).wrapping_add(1);
        seed >> 33
    };

    let words: Vec<String> = (0..2000)
        .map(|i| format!("Item-{i:04}-{}", i * 7919 % 1000))
        .collect();

    for row in 0..row_max {
        for col in 0..col_max {
            match mode {
                "varied" => {
                    if col % 2 == 1 {
                        let word = &words[(next() % 2000) as usize];
                        worksheet.write_string(row, col, word)?;
                    } else {
                        let number = (next() % 10_000_000) as f64 / 100.0;
                        worksheet.write_number(row, col, number)?;
                    }
                }
                _ => {
                    if col % 2 == 1 {
                        worksheet.write_string(row, col, "Foo")?;
                    } else {
                        worksheet.write_number(row, col, 12345.0)?;
                    }
                }
            }
        }
    }

    let levels: Vec<u8> = (1..=9).collect();
    let mut times: Vec<Vec<f64>> = vec![vec![]; levels.len()];
    let mut sizes = vec![0usize; levels.len()];

    // Round 0 is a warm-up and is discarded. Levels are interleaved within
    // each round so that machine drift affects all levels equally.
    for round in 0..=rounds {
        for (i, &level) in levels.iter().enumerate() {
            workbook.set_zip_compression_level(level);

            let start = Instant::now();
            let buffer = workbook.save_to_buffer()?;
            let elapsed = start.elapsed().as_secs_f64() * 1000.0;

            sizes[i] = buffer.len();
            if round > 0 {
                times[i].push(elapsed);
            }
        }
    }

    println!("mode={mode} rows={row_max} cols={col_max} rounds={rounds}");
    println!("level,size_bytes,size_vs_6,median_ms,min_ms,p10_ms,p90_ms,time_vs_6");

    let stats: Vec<(f64, f64, f64, f64)> = times
        .iter_mut()
        .map(|t| {
            t.sort_by(|a, b| a.partial_cmp(b).unwrap());
            let pct = |p: f64| t[((t.len() - 1) as f64 * p).round() as usize];
            (pct(0.5), t[0], pct(0.1), pct(0.9))
        })
        .collect();

    let (median_6, size_6) = (stats[5].0, sizes[5] as f64);
    for (i, &level) in levels.iter().enumerate() {
        let (median, min, p10, p90) = stats[i];
        println!(
            "{level},{},{:.3},{median:.1},{min:.1},{p10:.1},{p90:.1},{:.2}",
            sizes[i],
            sizes[i] as f64 / size_6,
            median / median_6
        );
    }

    Ok(())
}
