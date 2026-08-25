# Changelog

## [Unreleased]

## [2.1.1] - 2026-08-25

### Bug fixes

- Fixed a `StringIndexError` when strings terminating in a multi-byte character were passed from Excel to Julia, e.g. `JuliaCall("identity", "xxx£")`. Present since v0.2.17, the first version to use the wire format for the Excel -> Julia direction.

## [2.1.0] - 2026-08-24

### New features

- New `display_results` function: switch it on to echo, in the Julia console, both the expression or function call arriving from Excel (`question>`) and the value being returned to it (`answer>`) - useful for seeing exactly what `JuliaCall`/`JuliaEval` are doing. Example [here](README.md#demo).
- Clearer Julia console output when a `JuliaCall`/`JuliaEval` call fails: the error and the expression that caused it are now shown in colour, making them easier to spot among other console output.
- `JuliaExcel.last_question` and `JuliaExcel.last_answer` record the most recent expression or function call from Excel and the value returned to it, whichever of `JuliaCall`/`JuliaEval` was used. A new `JuliaExcel.answer_again()` re-evaluates `last_question`, and you can wrap it with `@enter` or your own debugging tools.
- New `JuliaExcel.serve_xl` function: attaches a Julia session that's already running (e.g. one open in VS Code) to Excel, instead of always launching a fresh process via `JuliaLaunch`. Call `JuliaExcel.serve_xl()` to attach to the single running Excel process automatically, or give the process id explicitly if more than one is running - get it from Excel's new `JuliaExcelPID()` worksheet function. Also adds `JuliaExcel.stop_server()` and `JuliaExcel.server_status()` for managing a session attached this way.
- `JuliaLaunch` and `JuliaCall`/`JuliaEval` are more resilient to the Julia session behind them changing or disappearing mid-session, e.g. after `JuliaExcel.stop_server()`, or a session attached via `JuliaExcel.serve_xl` on a different port.

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
| One-way (Excel to Julia): `JuliaCall("sum", vector of 100,000 Doubles)` | 1.45 s | 0.11 s | 0.11 s | 0.054 s |
| One-way (Julia to Excel): `JuliaEval("collect((1:100000).*pi)")` | 0.31 s | 0.26 s | 0.17 s | 0.041 s |
| One-way (Julia to Excel): `JuliaEval("(1:100000).*pi")` | 0.31 s | 0.26 s | 0.17 s | 0.014 s |

*All figures measured on the same machine: Intel Core Ultra 9 288V, 32GB RAM*

## [1.0.1] - 2026-08-14

### Changes

* Simplifies installation.
* Adds a compat entry for DataFrames, to meet registration requirements.

## [1.0.0] - 2026-08-13

### Changes

This version changes how JuliaExcel works, from file-based communication between Excel and Julia (versions 0.2.17 and earlier) to a local HTTP connection: an HTTP.jl server on the Julia side, called from Excel/VBA using a component built into Windows for making HTTP requests (MSXML2.XMLHTTP.6.0).

Combined with the changes already made in v0.2.17, this reduces latency and increases throughput between Excel and Julia:

| Test | v0.2.16 | v0.2.17 | v1.0 |
|---|---|---|---|
| Latency:</br>`JuliaEval("1+1")` | 7.2 ms | 6.4 ms | 1.1 ms |
| Two-way:</br>`JuliaCall("identity", vector of 100,000 Doubles)` | 1.56 s | 0.37 s | 0.27 s |
| One-way (Excel to Julia):</br>`JuliaCall("sum", vector of 100,000 Doubles)` | 1.45 s | 0.11 s | 0.11 s |
| One-way (Julia to Excel):</br>`JuliaEval("collect((1:100000).*pi)")` | 0.31 s | 0.26 s | 0.17 s |

These figures were measured on a Windows laptop (Intel Core Ultra 9 288V, 32GB RAM). To test on your PC: set the workbook to not be an add-in (e.g. from the VBE's Immediate window, `Workbooks("JuliaExcel.xlam").IsAddin = False`), then click the "Check Performance!" button - results are printed to the VBA Immediate window (Ctrl+G in the VBE).

Other changes in this release:

  * When a call made via `JuliaCall` or `JuliaCallVBA` raises an error, the arguments passed from Excel are now saved to `JuliaExcel.args_from_xl` before the call is attempted, and the error message printed in the Julia console includes a ready-made expression for reproducing the call. So you can debug the failure directly at the Julia REPL, without needing to repeat the call from Excel.

  * `JuliaLaunch` now adds `--threads=auto,1` to `CommandLineOptions` by default, giving Julia's HTTP server its own dedicated thread rather than sharing one with other work in the session.

  * Removes the restriction on how Windows Terminal is configured (#9).

  * Removes the VBA functions `JuliaResultFile` and `JuliaUnserialiseFile` and the previously-exported Julia functions `htd` and `hts`.

## [0.2.17] - 2026-08-07

### Changes

This version has better performance when passing large arrays of data between Julia and Excel. On a test laptop:

#### v0.2.16
`JuliaCall("sum",RANDARRAY(250000))` executes in 5900 milliseconds
`JuliaCall("sum",RANDARRAY(500000))` returns an error

#### v0.2.17
`JuliaCall("sum",RANDARRAY(250000))` executes in 330 milliseconds
`JuliaCall("sum",RANDARRAY(500000))` executes in 725 milliseconds

## [0.2.16] - 2025-12-28

### Changes

- **Improved Floating-Point Precision**
Communication of 32-bit and 64-bit floating-point numbers between Excel/VBA and Julia now preserves full precision.
*Example:* Previously, `=JuliaCall("identity",SQRT(2))-SQRT(2)` returned `4.88498E-15`; it now returns `0`.
- **Improved Array Support**
Multi-dimensional arrays (up to 8 dimensions) can now be passed between VBA and Julia. Earlier versions supported only 1D and 2D arrays.
- **Improved Dictionary Handling**
Dictionaries can now be passed in both directions: VBA ↔ Julia. Previously, only Julia → VBA was supported.
- **Regional Settings Compatibility**
JuliaExcel now works correctly when the Windows decimal symbol is not `.`.

## [0.2.15] - 2025-10-30

### Changes

* Communication between Julia and Excel (which is file-based) now has retry logic for better reliability, especially in tight loops or high-frequency calls.

## [0.2.12] - 2025-10-16

### Changes

- This version is compatible with Julia 1.12.
- Julia variables of unsupported type are now returned to Excel as their string representation, via `Base.show`.
- Improved handling of strings containing characters with high code points.

## [0.2.10] - 2023-12-03

### Changes

In prior versions, two functions (`GetCurrentProcessID` and `IsWindow`) were made visible as worksheet functions when they should not have been. This version corrects that.

## [0.2.9] - 2023-09-20

### Changes

Now launches Julia in interactive mode, i.e. with the `-i` command line option. The benefit is that when using [OhMyREPL](https://github.com/KristofferC/OhMyREPL.jl), `Ctrl + R` works correctly to see command history.

## [0.2.8] - 2022-07-12

### Changes

1) When [JuliaLaunch](https://github.com/PGS62/JuliaExcel.jl#julialaunch) is called, `julia.exe` must first be located: if it's on the PATH, the first one found there is used; otherwise, we search for the most recently modified copy of `julia.exe` in the locations where the Windows Julia installer places it. This change adds further such locations, in response to #10.
2) The value returned to Excel is no longer displayed in the Julia REPL.

## [0.2.7] - 2022-03-07

### Changes

If the Julia code being called from Excel throws an exception, then a stack trace of the error is displayed in the Julia window (uses [showerror](https://docs.julialang.org/en/v1/base/io-network/#Base.showerror)). Particularly helpful when the Julia code is under development.

## [0.2.6] - 2021-12-18

### Changes

Fixes a bug in VBA method JuliaEvalVBA.

## [0.2.5] - 2021-12-15

### Changes

Now compatible with Excel 2013

## [0.2.4] - 2021-12-10

### Changes

Now possible to run Julia on Ubuntu Linux running under Windows Subsystem for Linux.

The arguments to `JuliaLaunch` have changed in a not backwards-compatible way, so any existing calls to `JuliaLaunch` in workbooks or VBA code will need to be amended when upgrading to this version.

## [0.2.3] - 2021-12-02

### Changes

Dictionaries now handled.

If `JuliaEvalVBA` and `JuliaCallVBA` evaluate on the Julia side to `result` where `typeof(result) <: AbstractDict`, then in VBA the return from the function is of type `Scripting.Dictionary`.

## [0.2.2] - 2021-11-19

### Changes

Now launches Julia with the `--threads=auto` command line option.

## [0.2.1] - 2021-11-17

The first documented release.

## [0.1.0] - 2021-11-06

First release. No documentation yet!
