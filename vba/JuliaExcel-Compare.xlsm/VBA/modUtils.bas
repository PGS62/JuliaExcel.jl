Attribute VB_Name = "modUtils"
' Copyright (c) 2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme
Option Explicit

#If VBA7 And Win64 Then
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#Else
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'             Local copy of JuliaExcel.xlam's modUtils.ElapsedTime - duplicated rather than referenced
'             across VBA projects since this workbook has no VBA project reference to JuliaExcel.xlam.
' -----------------------------------------------------------------------------------------------------------------------
Function ElapsedTime() As Double
          Dim A As Currency
          Dim B As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter A
3         QueryPerformanceFrequency B
4         ElapsedTime = A / B

5         Exit Function
ErrHandler:
6         ReThrow "ElapsedTime", Err
End Function

'Local copy of JuliaExcel.xlam's modUtils.Throw.
Sub Throw(ByVal ErrorString As String)
1         If Len(ErrorString) > 32000 Then
2             Err.Raise vbObjectError + 1, , Left$(ErrorString, 1) & Right$(ErrorString, 31999)
3         Else
4             Err.Raise vbObjectError + 1, , Right$(ErrorString, 32000)
5         End If
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReThrow
' Purpose    : Common error handling to be used in the error handler of all methods. Local copy of
'              JuliaExcel.xlam's modUtils.ReThrow.
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

'Local copy of JuliaExcel.xlam's modMain.ThrowIfError.
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

Public Function IsInCollection(oColn As Object, Key As String) As Boolean
1         On Error GoTo ErrHandler
2         VarType (oColn(Key))
3         IsInCollection = True
4         Exit Function
ErrHandler:
End Function

Function GetEnvironmentVariable(Expression As String)
1     GetEnvironmentVariable = Environ(Expression)
End Function


Function IsRangeBlank(r As Range) As Boolean

1         IsRangeBlank = (Application.WorksheetFunction.CountA(r) = 0)

End Function



Sub ReleaseCleanup()
    On Error GoTo ErrHandler
    ClearOutSheets
    Application.Run "SolumAddin.xlam!StandardReleaseCleanup", ThisWorkbook

    Exit Sub
ErrHandler:
    ReThrow "ReleaseCleanup", Err
End Sub

