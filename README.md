# JuliaExcel

Call Julia from Excel spreadsheets and VBA.

## Installation

 * First ensure you have both [Julia](https://julialang.org/) and Microsoft Office installed. JuliaExcel works best with Office 365 or Office 2021, both of which support [dynamic array formulas](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).
 * Launch Julia and switch to the Package REPL with the `]` key.
 * Type `add https://github.com/PGS62/JuliaExcel.jl` then the `Enter` key. This installs the Julia code and downloads an installer for the associated Excel addin.
 * Switch back to the REPL with the `Backspace` key.
 * `using JuliaExcel` then the `Enter` key.
 * `JuliaExcel.installme()` then the `Enter` key. This installs the addin JuliaExcel.xlam to your Excel Addins folder.
 * Click through a couple of dialogs.
 * Launch Excel. Check that the JuliaExcel functions are available by typing `=Julia` into a worksheet cell and checking that the auto-complete offers `JuliaCall`, `JuliaEval`, `JuliaInclude` etc.

![installation](screenshots/installation.gif)

## Functions
JuliaExcel makes the following functions available from Excel worksheets and from VBA:

#### `JuliaLaunch`
Enter `=JuliaLaunch()` into a cell to launch an instance of Julia which will "listen" to the current Excel instance.

#### `JuliaInclude`
Load a Julia source file into the Julia process.

#### `JuliaEval`
Evaluate any Julia code, be it `1+1` or a call to a complex Julia function that you have loaded.

#### `JuliaCall`
Call a named Julia function, passing in arguments that reference ranges on the worksheet.

#### `JuliaSetVar`
Set a global variable in the Julia process to be equal to the contents of a range on the worksheet.

## Examples

## Alternatives

## How it works

## Shortcomings



Philip Swannell
6 November 2021
