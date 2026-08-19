# Changelog

## [2.0.0] - 2026-08-19

### Breaking changes

- Excel errors passed into Julia now decode to a new `ExcelError` type, not a `String`. Previously, an Excel error value (e.g. `#DIV/0!`) passed as an argument to `JuliaCall`/`JuliaEval` arrived in Julia as a string like `"#ExcelError2007!"`. It now arrives as `ExcelError(2007)` - a distinct type. This means:
  - A function like `identity` now correctly returns the *same* Excel error, round-tripping unchanged (previously it came back as literal text).
  - A function that isn't written to expect an error and tries to operate on one now fails immediately with a clear `MethodError`, rather than silently treating error text as ordinary string data.
  - Code relying on the old string representation will need updating.
- Dictionaries passed from Excel now decode to a concretely-typed `Dict{K,V}` when every key and every value share a type, rather than always `Dict{Any,Any}`. Code dispatching narrowly on `Dict{Any,Any}` should broaden to `Dict`/`AbstractDict`.

### New features

- Large numeric arrays now transfer substantially faster between Excel and Julia, in both directions - every number is packed into a fixed number of characters with no per-element bookkeeping, and the core conversion between numbers and that packed form happens as one bulk operation on both sides, rather than looping over each number individually.
- Friendlier `JuliaCall`/`JuliaEval` error messages: an expression that errors now returns a short summary to Excel (e.g. `MethodError: no method matching +(::Int64, ::String)...`) instead of a raw exception dump like `#(MethodError(+, (1, "1"), 0x000000000000979f))!`. The full error and stacktrace are printed to the Julia console.
- Julia `Range`s (`UnitRange`, `StepRange` etc.) now transfer to Excel far faster regardless of size - only the first value, step, and length cross the wire, and VBA reconstructs the array from those.
- `Byte` support: VBA's `Byte` type now round-trips to Julia's `UInt8` (previously unsupported in both directions).

### Performance

| Test | v0.2.16 | v0.2.17 | v1.0 | v2.0 |
|------|---------|---------|------|------|
| Latency: `JuliaEval("1+1")` | 7.2 ms | 6.4 ms | 1.1 ms | 1.2 ms |
| Two-way: `JuliaCall("identity", vector of 100,000 Doubles)` | 1.56 s | 0.37 s | 0.27 s | 0.10 s |
| One-way (Excel to Julia): `JuliaCall("sum", vector of 100,000 Doubles)` | 1.45 s | 0.11 s | 0.11 s | 0.05 s |
| One-way (Julia to Excel): `JuliaEval("collect((1:100000).*pi)")` | 0.31 s | 0.26 s | 0.17 s | 0.05 s |
| One-way (Julia to Excel): `JuliaEval("(1:100000).*pi")` | 0.31 s | 0.26 s | 0.17 s | 0.01 s |

*All figures measured on the same machine: Intel Core Ultra 9 288V, 32GB RAM*
