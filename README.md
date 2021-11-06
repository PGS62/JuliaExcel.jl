# JuliaVBA

Call [Julia](https://julialang.org/) from Excel spreadsheets and VBA.

## Installation

 * Assumes you have Julia and Microsoft Office installed. JuliaVBA works best with Office 365 or Office 2021, both of which [dynamic array formulas](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).
 * Launch Julia and switch to the Package REPL with the `]` key.
 * Type `add https://github.com/PGS62/JuliaVBA.jl` then the `Enter` key. This installs the Julia code and downloads an installer for the associated Excel addin.
 * Switch back to the REPL with the `Backspace` key.
 * `using JuliaVBA` then the `Enter` key.
 * `JuliaVBA.installme()` then the `Enter` key. This installs the addin JuliaVBA.xlam to your Excel Addins folder.
 * Click through a couple of dialogs.
 * Launch Excel. Check that the JuliaVBA functions are available by typing `=Julia` into a worksheet cell and checking that the auto-complete offers `JuliaCall`, `JuliaEval`, `JuliaInclude` etc.

## Features

## Functions

## Examples

## Alternatives

## How it works

## Shortcomings



Philip Swannell
6 November 2021
