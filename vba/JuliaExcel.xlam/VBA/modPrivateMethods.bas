Attribute VB_Name = "modPrivateMethods"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

' Home for procedures that need to be visible across modules (so, not procedure-level Private) but
' must not be discoverable from a worksheet formula bar - Option Private Module hides everything
' here from autocomplete, the Insert Function dialog and the Macro list. Two kinds of procedure end
' up here: widely-used internal helpers (ThrowIfError, ReThrow) called from many other modules, and
' otherwise-Private-in-spirit procedures that a test living in a different module needs to call
' directly (ConcatenateExpressions, tested by TestConcatenateExpressions in modTest.bas).

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
' -----------------------------------------------------------------------------------------------------------------------
Function ThrowIfError(Data As Variant)
1         ThrowIfError = Data
2         If VarType(Data) = vbString Then
3             If Left$(Data, 1) = "#" Then
4                 If Right$(Data, 1) = "!" Then
5                     Throw CStr(Data)
6                 End If
7             End If
8         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReThrow
' Purpose    : Common error handling to be used in the error handler of all methods.
' Parameters :
'  FunctionName: The name of the function from which ReThrow is called, typically in the function's error handler.
'  Error       : Err, the error object.
'  ReturnString: Pass in True if the method is a "top level" method that's exposed to the user and we wish for the
'                function to return an error string (starts with #, ends with !).
'                Pass in False if we want to (re)throw an error, with annotated Description.
' -----------------------------------------------------------------------------------------------------------------------
Function ReThrow(FunctionName As String, Error As ErrObject, Optional ReturnString As Boolean = False)
          Dim ErrorDescription As String
          Dim ErrorNumber As Long
          Dim LineDescription As String

1         ErrorDescription = Error.Description
2         ErrorNumber = Err.Number

          'Build up call stack, i.e. annotate error description by prepending #<FunctionName> and appending !
3         If Erl = 0 Then
4             LineDescription = " (line unknown): "
5         Else
6             LineDescription = " (line " & CStr(Erl) & "): "
7         End If
8         ErrorDescription = "#" & FunctionName & LineDescription & ErrorDescription & "!"

9         If ReturnString Then
10            ReThrow = ErrorDescription
11        Else
12            Err.Raise ErrorNumber, , ErrorDescription
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConcatenateExpressions
' Purpose    : It's convenient to be able to pass in a multi-line expression, which we first concatenate with semi-colon
'              delimiter before passing to Julia for evaluation
' -----------------------------------------------------------------------------------------------------------------------
Function ConcatenateExpressions(JuliaExpression As Variant) As String
          Dim i As Long
          Dim NC As Long
          Dim Tmp() As String
1         On Error GoTo ErrHandler
2         If TypeName(JuliaExpression) = "Range" Then
3             JuliaExpression = JuliaExpression.Value
4         End If
5         Select Case NumDimensions(JuliaExpression)
              Case 0
6                 ConcatenateExpressions = CStr(JuliaExpression)
7             Case 1
8                 ConcatenateExpressions = VBA.Join(JuliaExpression, ";")
9             Case 2
10                NC = UBound(JuliaExpression, 2) - LBound(JuliaExpression, 1) + 1
11                If NC > 1 Then Throw "When passed as an array or a Range, JuliaExpression should have only one column, but got " + CStr(NC) + " columns"
12                ReDim Tmp(LBound(JuliaExpression, 1) To UBound(JuliaExpression, 1))
13                For i = LBound(Tmp) To UBound(Tmp)
14                    Tmp(i) = JuliaExpression(i, LBound(JuliaExpression, 2))
15                Next
16                ConcatenateExpressions = VBA.Join(Tmp, ";")
17            Case Else
18                Throw "Too many dimensions in JuliaExpression"
19        End Select
20        Exit Function
ErrHandler:
21        ReThrow "ConcatenateExpressions", Err
End Function

