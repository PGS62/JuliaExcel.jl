Attribute VB_Name = "modPostMessage"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme
Option Explicit
Option Private Module

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, _
                                      ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                      ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
                                     (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                    (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
#Else
    Private Declare  Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare  Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                      ByVal lpWindowName As String) As Long
    Private Declare  Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
                                     (ByVal hwnd As Long) As Long
    Private Declare  Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare  Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
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
