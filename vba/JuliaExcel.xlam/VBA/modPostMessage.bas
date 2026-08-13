Attribute VB_Name = "modPostMessage"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module
Private Const GW_HWNDNEXT = 2

#If VBA7 And Win64 Then
Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" _
    (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" _
    (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32.dll" _
    (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" _
    (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : NumWindowsWithCaption
' Purpose    : How many windows are open whose caption includes sCaption?
' -----------------------------------------------------------------------------------------------------------------------
Function NumWindowsWithCaption(ByVal sCaption As String) As Long

          'https://docs.microsoft.com/en-gb/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN
          Const GW_HWNDNEXT = 2

          Dim lhWndP As LongPtr
          Dim sStr As String
1         On Error GoTo ErrHandler
2         NumWindowsWithCaption = False
3         lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
4         Do While lhWndP <> 0
5             sStr = WindowTitleFromHandle(lhWndP)
6             If InStr(1, sStr, sCaption) > 0 Then
7                 NumWindowsWithCaption = NumWindowsWithCaption + 1
8             End If
9             lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
10        Loop

11        Exit Function
ErrHandler:
12        ReThrow "NumWindowsWithCaption", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetHandleFromPartialCaption
' Purpose    : Get a window handle for a window whose title contains the string sCaption
' Adapted from
' https://stackoverflow.com/questions/25098263/how-to-use-findwindow-to-find-a-visible-or-invisible-window-with-a-partial-name
' -----------------------------------------------------------------------------------------------------------------------
Function GetHandleFromPartialCaption(ByRef lwnd As LongPtr, ByVal sCaption As String) As Boolean

          'https://docs.microsoft.com/en-gb/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN
          Const GW_HWNDNEXT = 2

          Dim lhWndP As LongPtr
          Dim sStr As String
1         On Error GoTo ErrHandler
2         GetHandleFromPartialCaption = False
3         lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
4         Do While lhWndP <> 0
5             sStr = WindowTitleFromHandle(lhWndP)
6             If InStr(1, sStr, sCaption) > 0 Then
7                 GetHandleFromPartialCaption = True
8                 lwnd = lhWndP
9                 Exit Do
10            End If
11            lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
12        Loop

13        Exit Function
ErrHandler:
14        ReThrow "GetHandleFromPartialCaption", Err
End Function

Function WindowTitleFromHandle(lhWndP As LongPtr)
          Dim sStr As String
1         On Error GoTo ErrHandler
2         sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
3         GetWindowText lhWndP, sStr, Len(sStr)
4         sStr = Left$(sStr, Len(sStr) - 1)
5         WindowTitleFromHandle = sStr
6         Exit Function
ErrHandler:
7         ReThrow "WindowTitleFromHandle", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsFunctionWizardActive
' Purpose    : Tests if the Excel Function Wizard is in use.
'            : See discussion at https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
' -----------------------------------------------------------------------------------------------------------------------
Function IsFunctionWizardActive() As Boolean

          Dim ExcelPID As Long
          Dim lhWndP As LongPtr
          Dim WindowPID As Long
          Dim WindowTitle As String
          Const FunctionWizardCaption = "Function Arguments" 'This won't work for non English-language Excel
          
1         On Error GoTo ErrHandler
2         If TypeName(Application.Caller) = "Range" Then
              'The "CommandBars test" below is usually sufficient to determine that the Function Wizard is active,
              'but can sometimes give a false positive. Example: When a csv file is opened (via File Open) then all
              'active workbooks are calculated (even if calculation is set to manual!) with
              'Application.CommandBars("Standard").Controls(1).Enabled being False.
              'So apply a further test using Windows API to loop over all windows checking for a window with title
              '"Function  Arguments", checking also the process id.
3             If Not Application.CommandBars("Standard").Controls(1).Enabled Then
4                 ExcelPID = GetCurrentProcessId()
5                 lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
6                 Do While lhWndP <> 0
7                     WindowTitle = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
8                     GetWindowText lhWndP, WindowTitle, Len(WindowTitle)
9                     WindowTitle = Left$(WindowTitle, Len(WindowTitle) - 1)
10                    If WindowTitle = FunctionWizardCaption Then
11                        GetWindowThreadProcessId lhWndP, WindowPID
12                        If WindowPID = ExcelPID Then
13                            IsFunctionWizardActive = True
14                            Exit Function
15                        End If
16                    End If
17                    lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
18                Loop
19            End If
20        End If

21        Exit Function
ErrHandler:
22        ReThrow "IsFunctionWizardActive", Err
End Function

