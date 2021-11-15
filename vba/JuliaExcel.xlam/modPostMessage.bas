Attribute VB_Name = "modPostMessage"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme
Option Explicit
Option Private Module
Private Const GW_HWNDNEXT = 2

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hWnd As LongPtr, _
        ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As Long
    Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" _
        (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" _
        (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32.dll" _
        (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
#Else
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
        (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
        (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" _
        (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PostMessageToJulia
' Purpose    : Sends a command to a running Julia process. Faster and more robust than using Application.SendKeys
' Remarks    : Figuring out the correct lParam argument to pass to PostMessage is tricky, Spy++ is helpful, and is
'              installed on my PC at:
'              C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\Tools\spyxx_amd64.exe
'              More references at:
'              https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessagea
'              https://www.codeproject.com/Tips/1029254/SendMessage-and-PostMessage
'              Key codes are at:
'              https://docs.microsoft.com/en-gb/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN
' Parameters :
'  HwndJulia : Window handle for the Julia REPL
' -----------------------------------------------------------------------------------------------------------------------
Sub PostMessageToJulia(HwndJulia As LongPtr)
          
          Dim i As Long
            
          'https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-char
          Const WM_CHAR = &H102
            
          'In case there's some random text at the Julia REPL, send {ESCAPE}{BACKSPACE} three times
1         For i = 1 To 3
2             PostMessage HwndJulia, WM_CHAR, ByVal 27, ByVal &H10001
3             PostMessage HwndJulia, WM_CHAR, ByVal 8, ByVal &H10001
4         Next i
          'One more {BACKSPACE} should be enough to switch Julia out of Package REPL mode if it's in it.
5         PostMessage HwndJulia, WM_CHAR, ByVal 8, ByVal &H10001

          'Send "srv_xl". In this block the lParam arg is not as per Spy++, but I _
           think it's enough that the first bit is set to 1 to indicate that the _
           key should be repeated only once. _
           https://docs.microsoft.com/en-gb/windows/win32/inputdev/wm-char?redirectedfrom=MSDN
6         PostMessage HwndJulia, WM_CHAR, ByVal Asc("s"), ByVal &H1
7         PostMessage HwndJulia, WM_CHAR, ByVal Asc("r"), ByVal &H1
8         PostMessage HwndJulia, WM_CHAR, ByVal Asc("v"), ByVal &H1
9         PostMessage HwndJulia, WM_CHAR, ByVal Asc("_"), ByVal &H1
10        PostMessage HwndJulia, WM_CHAR, ByVal Asc("x"), ByVal &H1
11        PostMessage HwndJulia, WM_CHAR, ByVal Asc("l"), ByVal &H1

          'Send "(){Enter}"
12        PostMessage HwndJulia, WM_CHAR, ByVal Asc("("), ByVal &HA0001
13        PostMessage HwndJulia, WM_CHAR, ByVal Asc(")"), ByVal &HB0001
14        PostMessage HwndJulia, WM_CHAR, ByVal Asc(vbLf), ByVal &H1C0001

15        Exit Sub
ErrHandler:
16        Throw "#PostMessageToJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

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
14        Throw "#GetHandleFromPartialCaption (line " & CStr(Erl) + "): " & Err.Description & "!"
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
7         Throw "#WindowTitleFromHandle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsFunctionWizardActive
' Purpose    : Tests if the Excel Function Wizard is in use.
'            :See discussion at https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
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
22        Throw "#IsFunctionWizardActive (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

