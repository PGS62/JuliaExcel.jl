# JuliaExcel

Call Julia from Excel spreadsheets and VBA.

## Installation

 * Both [Julia](https://julialang.org/) and Microsoft Office must be installed on your PC, with Excel not running.
 * Launch Julia and switch to the Package REPL with the `]` key.
 * Type `add https://github.com/PGS62/JuliaExcel.jl` followed by the `Enter` key.
 * Type `using JuliaExcel` followed by `Enter`.
 * Type `JuliaExcel.installme()` followed by `Enter`.
 * Click through a couple of dialogs.
 * Launch Excel. Check that the JuliaExcel functions are available by typing `=Julia` into a worksheet cell and checking that the auto-complete offers `JuliaCall`, `JuliaEval`, `JuliaInclude` etc.

![installation](images/installation.gif)

## Functions
JuliaExcel makes the following functions available from Excel worksheets and from VBA:

#### `JuliaLaunch`
Enter `=JuliaLaunch()` into a cell to launch a JUlia process that "listens" to the current Excel process.

#### `JuliaInclude`
Load a Julia source file into the Julia process.

#### `JuliaEval`
Evaluate any Julia code, be it `1+1` or a call to any Julia function that you have loaded.

#### `JuliaCall`
Call a named Julia function, passing in arguments that reference ranges on the worksheet.

#### `JuliaSetVar`
Set a global variable in the Julia process to be equal to the contents of a range on the worksheet.

## Demo
![demo2](images/Demo4.gif)
## Alternatives

## How it works

## Shortcomings



Philip Swannell
6 November 2021
