Attribute VB_Name = "modRegister"
Option Explicit
Option Private Module

Sub RegisterAll()
    Application.ScreenUpdating = False
    'Without setting .IsAddin to False, I get errors: "Cannot edit a macro on a hidden workbook. Unhide the workbook using the Unhide command."
    ThisWorkbook.IsAddin = False
    RegisterJuliaInclude
    RegisterJuliaEval
    RegisterJuliaCall
    RegisterJuliaCall2
    RegisterJuliaSetVar
    ThisWorkbook.IsAddin = True
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaLaunch
' Purpose    : Register the function JuliaLaunch with the Excel function wizard, to be called from the WorkBook_Open
'              event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaLaunch()
    Const Description As String = "Launches a local Julia session which ""listens"" to the current Excel session " & _
                                  "and responds to calls to JuliaEval etc.."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 2)
    ArgDescs(1) = "If TRUE, then the Julia session window is minimised, if FALSE (the default) then the window is " & _
                  "sized normally."
    ArgDescs(2) = "The location of julia.exe. If omitted, then the function searches for julia.exe, first on the " & _
                  "path and then at the default locations for Julia installation on Windows, taking the most " & _
                  "recently installed version if more than one is available."
    Application.MacroOptions "JuliaLaunch", Description, , , , , "JuliaExcel", , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function JuliaLaunch failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaInclude
' Purpose    : Register the function JuliaInclude with the Excel function wizard, to be called from the WorkBook_Open
'              event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaInclude()
    Const Description As String = "Load a Julia source file into the Julia process, with the likely intention of " & _
                                  "making additional functions available via JuliaEval and JuliaCall."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 2)
    ArgDescs(1) = "The full name of the file to be included."
    ArgDescs(2) = "Provides control over worksheet calculation dependency. Enter a cell or range that must be " & _
                  "calculated before JuliaInclude is executed."
    Application.MacroOptions "JuliaInclude", Description, , , , , "JuliaExcel", , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function JuliaInclude failed with error: " + Err.Description
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaEval
' Purpose    : Register the function JuliaEval with the Excel function wizard, to be called from the WorkBook_Open
'              event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaEval()
    Const Description As String = "Evaluate a Julia expression and return the result to Excel or VBA."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 2)
    ArgDescs(1) = "Any valid julia code, as a string. Can also be a one-column range to evaluate multiple julia " & _
                  "statements."
    ArgDescs(2) = "Provides control over worksheet calculation dependency. Enter a cell or range that must be " & _
                  "calculated before JuliaEval is executed."
    Application.MacroOptions "JuliaEval", Description, , , , , "JuliaExcel", , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function JuliaEval failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaCall
' Purpose    : Register the function JuliaCall with the Excel function wizard, to be called from the WorkBook_Open
'              event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaCall()
    Const Description As String = "Call a named Julia function, passing in data from the worksheet or from VBA."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 2)
    ArgDescs(1) = "The name of a Julia function that's defined in the Julia session, perhaps as a result of prior " & _
                  "calls to JuliaInclude."
    ArgDescs(2) = "Zero or more arguments, which may be Excel ranges or variables in VBA code."
    Application.MacroOptions "JuliaCall", Description, , , , , "JuliaExcel", , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function JuliaCall failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaCall2
' Purpose    : Register the function JuliaCall2 with the Excel function wizard, to be called from the WorkBook_Open
'              event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaCall2()
    Const Description As String = "Call a named Julia function, passing in data from the worksheet or from VBA, " & _
                                  "with control of worksheet calculation dependency."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 3)
    ArgDescs(1) = "The name of a Julia function that's available in the Main module of the running Julia session."
    ArgDescs(2) = "Provides control over worksheet calculation dependency. Enter a cell or range that must be " & _
                  "calculated before JuliaCall2 is executed."
    ArgDescs(3) = "Zero or more arguments, such as Excel ranges or nested formulas."
    Application.MacroOptions "JuliaCall2", Description, , , , , "JuliaExcel", , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function JuliaCall2 failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaSetVar
' Purpose    : Register the function JuliaSetVar with the Excel function wizard, to be called from the WorkBook_Open
'              event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaSetVar()
    Const Description As String = "Set a global variable in the Julia process."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 3)
    ArgDescs(1) = "The name of the variable to be set. Must follow Julia's rules for allowed variable names."
    ArgDescs(2) = "An Excel range (from which the .Value2 property is read) or more generally a number, string, " & _
                  "Boolean, Empty or array of such types. When called from VBA, nested arrays are supported."
    ArgDescs(3) = "Provides control over worksheet calculation dependency. Enter a cell or range that must be " & _
                  "calculated before JuliaSetVar is executed."
    Application.MacroOptions "JuliaSetVar", Description, , , , , "JuliaExcel", , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function JuliaSetVar failed with error: " + Err.Description
End Sub



