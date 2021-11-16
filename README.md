# JuliaExcel

Call Julia functions from Microsoft Excel worksheets and from VBA.  

Compatible with Excel's [dynamic array formulas](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).

Windows only.

## Contents
[Installation](#installation)  
[Functions](#functions)  
[Demo](#demo)  
[Example VBA](#example-vba)  
[Function Documentation](#function-documentation)  
[Marshalling](#marshalling)  
[Alternatives](#alternatives)  
[How JuliaExcel works](#how-juliaexcel-works)  
[Shortcomings](#shortcomings)  

## Installation
Installation does not require admin rights on the PC.
 * Both [Julia](https://julialang.org/) and Microsoft Office must be installed on your PC, with Excel not running.
 * Launch Julia and switch to the Package REPL with the `]` key.
 * Type `add https://github.com/PGS62/JuliaExcel.jl` followed by the `Enter` key.
 * Exit the Package REPL with the `Backspace` key, then type `using JuliaExcel` followed by `Enter`.
 * Type `JuliaExcel.installme()` followed by `Enter`.
 * Click through a couple of dialogs.
 * Launch Excel. Check that the JuliaExcel functions are available by typing `=Julia` into a worksheet cell and checking that the auto-complete offers `JuliaCall`, `JuliaEval`, `JuliaInclude` etc.

The process is illustrated in the GIF below. F5 to replay.

![installation](images/installation.gif)

## Functions
JuliaExcel makes the following functions available from Excel worksheets and from VBA:

|Name|Description|
|----|-----------|
|[JuliaLaunch](#julialaunch)|Launches a local Julia session which "listens" to the current Excel session and responds to calls to JuliaEval etc..|
|[JuliaInclude](#juliainclude)|Load a Julia source file into the Julia process, to make additional functions available via JuliaEval and JuliaCall.|
|[JuliaEval](#juliaeval)|Evaluate a Julia expression and return the result to an Excel worksheet.|
|[JuliaEvalFromVBA](#juliaevalfromvba)|Evaluate a Julia expression and return the result to VBA. Tuned for use from VBA rather than a worksheet.|
|[JuliaCall](#juliacall)|Call a named Julia function, passing in data from the worksheet.|
|[JuliaCallFromVBA](#juliacallfromvba)|Call a named Julia function from VBA code. Tuned for use from VBA rather than a worksheet.|
|[JuliaSetVar](#juliasetvar)|Set a global variable in the Julia process.|

## Demo
Here's a quick demonstration of the functions in action.
 * See how the Julia session on the left responds to the action in Excel on the right.
 * The annotations in brown text ("Formula at...") are to make the what's happening in the demo clearer. They won't appear when you try JuliaExcel for yourself!
 * You can replay the GIF by hitting F5.
![demo2](images/Demo4.gif)

## Example VBA
The VBA code below makes a call to `JuliaLaunch` and `JuliaEval` and then pastes the result to range A1:J10 in a new worksheet. To run it, make sure that the project has a reference to JuliaExcel (VBA editor, Tools menu -> References).

```vba
Sub DemoCallFromVBA()

    Dim ResultFromJulia As Variant, PasteHere As Range
    
    JuliaLaunch
    
    ResultFromJulia = JuliaEval("(1:10).^(1:10)'")

    Set PasteHere = Application.Workbooks.Add.Worksheets(1) _
        .Cells(1, 1).Resize(UBound(ResultFromJulia, 1), _
        UBound(ResultFromJulia, 2))
    
    PasteHere.Value = ResultFromJulia

End Sub
```

## Function Documentation

#### _JuliaLaunch_
Launches a local Julia session which "listens" to the current Excel session and responds to calls to `JuliaEval` etc..
```vba
Function JuliaLaunch(Optional MinimiseWindow As Boolean, Optional ByVal JuliaExe As String)
```

|Argument|Description|
|:-------|:----------|
|`MinimiseWindow`|If TRUE, then the Julia session window is minimised, if FALSE (the default) then the window is sized normally.|
|`JuliaExe`|The location of julia.exe. If omitted, then the function searches for julia.exe, first on the path and then at the default locations for Julia installation on Windows, taking the most recently installed version if more than one is available.|


#### _JuliaInclude_
Load a Julia source file into the Julia process, to make additional functions available via `JuliaEval` and `JuliaCall`.
```vba
Function JuliaInclude(FileName As String)
```

|Argument|Description|
|:-------|:----------|
|`FileName`|The full name of the file to be included.|


#### _JuliaEval_
Evaluate a Julia expression and return the result to an Excel worksheet.
```vba
Function JuliaEval(ByVal JuliaExpression As Variant)
```

|Argument|Description|
|:-------|:----------|
|`JuliaExpression`|Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple Julia statements.|


#### _JuliaEvalFromVBA_
Evaluate a Julia expression and return the result to VBA. Designed for use from VBA rather than a worksheet and differs from `JuliaEval` in handling of 1-dimensional arrays, nested arrays and strings longer than 32,767 characters.
```vba
Function JuliaEvalFromVBA(ByVal JuliaExpression As Variant)
```

|Argument|Description|
|:-------|:----------|
|`JuliaExpression`|Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple Julia statements.|


#### _JuliaCall_
Call a named Julia function, passing in data from the worksheet.
```vba
Function JuliaCall(JuliaFunction As String, ParamArray Args())
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function that's defined in the Julia session, perhaps as a result of prior calls to `JuliaInclude`.|
|`Args...`|Zero or more arguments. Each argument may be a number, string, Boolean value, empty cell, an array of such values or an Excel range.|


#### _JuliaCallFromVBA_
Call a named Julia function from VBA code. Designed for use from VBA rather than a worksheet and differs from `JuliaCall` in handling of 1-dimensional arrays, nested arrays and strings longer than 32,767 characters.
```vba
Function JuliaCallFromVBA(JuliaFunction As String, ParamArray Args())
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function that's defined in the Julia session, perhaps as a result of prior calls to `JuliaInclude`.|
|`Args...`|Zero or more arguments. Each argument may be a number, string, Boolean value, empty cell, an array of such values or an Excel range.|


#### _JuliaSetVar_
Set a global variable in the Julia process.
```vba
Function JuliaSetVar(VariableName As String, RefersTo As Variant)
```

|Argument|Description|
|:-------|:----------|
|`VariableName`|The name of the variable to be set. Must follow Julia's [rules](https://docs.julialang.org/en/v1/manual/variables/#Allowed-Variable-Names) for allowed variable names.|
|`RefersTo`|An Excel range (from which the .Value2 property is read) or more generally a number, string, Boolean, Empty or array of such types. When called from VBA, nested arrays are supported.|

## Marshalling
Two question arose during implementation:

First, when data from a worksheet (or a VBA variable) is passed to `JuliaCall` or `JuliaSetVar`, that data is marshalled over to Julia. As what Julia type should the data arrive? Mostly, this is easy to decide, but what about one-dimensional arrays (from VBA) or ranges with just one column or one just row from an Excel worksheet? Should these have one-dimension or two over in Julia?

Second, after Julia has evaluated the expression, how should the result be marshalled in the opposite direction, back to Excel? Again this is easy to decide for scalars and two dimensional arrays, but what about for vectors in Julia?

There were three objectives to the design of the marshalling processes:
 1) Round-tripping should work, i.e. the formula `=JuliaCall("identity",x)` should return an identical copy of `x`, whatever the "shape" of `x`.
 2) Matrix arithmetic should work naturally. In Julia, the `*` operator does matrix multiplication, so marshalling should be such that the formula `=JuliaCall("*",Range1,Range2)` performs the same matrix
 multiplication as the formula `=MMULT(Range1,Range2`), which calls Excel's built-in `MMULT`.
 3) To allow use from `JuliaCall` of Julia's dot syntax for function broadcasting.
 
 The following marshalling scheme achieves the objectives:

 * Scalar values in Excel marshal back and forth to Julia as scalar values.
 * Two-dimensional arrays (or ranges) with more than one row and more than one column marshal back and forth as two dimensional.
 * Single-column ranges, when passed to `JuliaCall` or `JuliaSetVar`, arrive in Julia as vectors.
 * Conversely, if the result of an evaluation in Julia is a vector, then the return from 
 `JuliaCall` or `JuliaEval` is a two dimensional array with one column, which occupies a single column range on the worksheet.
 * Single-row ranges, when passed to `JuliaCall` or `JuliaSetVar`, arrive in Julia as 2-dimensional arrays with a single row.

 Click the black triangles below to see illustrations.
 
 <details><summary><u>Round-tripping of vectors and matrices</u></summary>
 <p>
  
 ![roundtripping](images/roundtripping.gif)
</p>
</details>

<details><summary>Matrix arithmetic</summary><p>

 ![matrixarithmetic](images/matrixarithmetic.gif)
</p></details>

<details><summary>Function broadcasting</summary><p>

 ![functionbroadcasting](images/functionbroadcasting.gif)
 </p></details>
  
## Alternatives
There is one alternative method of calling Julia from Excel of which I am aware:  

https://github.com/JuliaComputing/JuliaInXL.jl  

JuliaInXL has recently (October 2021) been made open source, having previously required a licence for commercial use. At the time of writing, it's not possible to call JuliaInXL from VBA and it is not compatible with dynamic array formulas when called from Excel worksheets. 

## Compatibility
JuliaExcel has been tested on Excel within Microsoft 365, both 32-bit and 64-bit. It _should_ work on earlier versions of Excel (perhaps back to Excel 2010) but it has not been tested on them.

## How JuliaExcel works
The implementation of JuliaExcel is very "low-tech". When a `JuliaEval` is called from a worksheet, the following happens:
1) VBA code (in JuliaExcel.xlam) writes the expression to a file in the JuliaExcel sub-folder of the temporary folder.
2) VBA code then uses the Windows API PostMessage to send keystrokes to the Julia window, the keystrokes are `srv_xl()`
3) That causes the Julia function in `srv_xl` (defined in JuliaExcel.jl) to execute. The function reads the expression file, evaluates it and writes to a result file.
4) The VBA code (in a wait loop since step 1) detects that the result file has been written, and unserialises the contents of the result file.

Other points to note:
 * `JuliaCall` is simply a wrapper to JuliaEval, with the arguments to `JuliaCall` being encoded using Julia's syntax for array literals.
 * The result file is written in a custom format designed to be fast to unserialise
 * There is obvious scope to improve this implementation by switching away from a file-based messaging system to to one based on sockets. Perhaps in a future version.

## Viewing the VBA code
The VBA project is password protected to prevent accidental changes. You can see the code [here](https://github.com/PGS62/JuliaExcel.jl/blob/master/vba/JuliaExcel.xlam/modMain.bas), or view it in the JuliaExcel.xlam by unprotecting with the password "JuliaExcel".

## Shortcomings



Philip Swannell
15 November 2021
