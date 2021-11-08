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

|Name|Description|
|----|-----------|
|JuliaLaunch|Launches a local Julia session which "listens" to the current Excel session and responds to calls to JuliaEval etc..|
|JuliaInclude|Load a Julia source file into the Julia process, to make additional functions available via JuliaEval and JuliaCall.|
|JuliaEval|Evaluate a Julia expression and return the result to Excel or VBA.|
|JuliaCall|Call a named Julia function, passing in data from the worksheet or from VBA.|
|JuliaCall2|Call a named Julia function, passing in data from the worksheet or from VBA, with control of worksheet calculation dependency.|
|JuliaSetVar|Set a global variable in the Julia process.|


## Demo
Here's a quick demonstration of the functions in action. Notice how the Julia session responds to the action over in Excel. Refresh your browser (F5) to restart.
![demo2](images/Demo4.gif)
## Alternatives

## How it works

## Shortcomings



Philip Swannell
6 November 2021
