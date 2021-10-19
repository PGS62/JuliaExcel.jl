Attribute VB_Name = "modSendMessage"
Option Explicit

'Declarations taken from https://jkp-ads.com/articles/apideclarations.asp
                                  
Private Declare PtrSafe Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hWnd As LongPtr, _
                                  ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As Long
                                  
                                  
Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr

'Source: https://renenyffenegger.ch/notes/development/languages/VBA/Win-API/index
Private Const WM_CHAR = &H102
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

'Source: https://docs.microsoft.com/en-gb/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN
Private Const VK_BACK = &H8             'BACKSPACE key
Private Const VK_RETURN = &HD           'ENTER key
Private Const VK_SHIFT = &H10           'SHIFT key
Private Const VK_ESCAPE = &H1B          'ESC key
Private Const VK_0 = &H30               '0 key
Private Const VK_9 = &H39               '9 key
Private Const VK_Z = &H5A               'Z key

Private Const GW_HWNDNEXT = 2


'References:
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessagea
'https://www.codeproject.com/Tips/1029254/SendMessage-and-PostMessage
'Key codes are at:
'https://docs.microsoft.com/en-gb/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN

'Spy++ is installed on my PC at:
'C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\Tools\spyxx_amd64.exe

'Turns out I can't spy on the Julia REPL as it's a Console Window
'https://stackoverflow.com/questions/37057816/why-does-spy-fail-with-console-windows

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SendMessageToJulia
' Author     : Philip Swannell
' Date       : 19-Oct-2021
' Purpose    : Sends keystrokes to Julia, much faster and more robust than using Application.SendKeys, but painful to
'              write the code, for which Spy++ is essential.
' Parameters :
'  HwndJulia:
' -----------------------------------------------------------------------------------------------------------------------
Sub SendMessageToJulia(HwndJulia As LongPtr)
    
    Dim i As Long
      
    'In case there's some random text at the Julia REPL...
    For i = 1 To 3
        'Send ESCAPE
        PostMessage HwndJulia, WM_KEYDOWN, VK_ESCAPE, &H10001
        PostMessage HwndJulia, WM_CHAR, ByVal 27, ByVal &H10001
        PostMessage HwndJulia, WM_KEYUP, VK_ESCAPE, &HC0010001

        'Send BACKSPCE
        PostMessage HwndJulia, WM_KEYDOWN, VK_BACK, &H10001
        PostMessage HwndJulia, WM_CHAR, ByVal 8, ByVal &H10001
        PostMessage HwndJulia, WM_KEYUP, VK_BACK, &HC0010001
    Next i

    'Send z
    PostMessage HwndJulia, WM_KEYDOWN, VK_Z, &H2C0001
    PostMessage HwndJulia, WM_CHAR, ByVal Asc("z"), ByVal &H2C0001
    PostMessage HwndJulia, WM_KEYUP, VK_Z, &HC02C0001

    'Send (
    PostMessage HwndJulia, WM_KEYDOWN, VK_SHIFT, &H360001
    PostMessage HwndJulia, WM_KEYDOWN, VK_9, &HA0001
    PostMessage HwndJulia, WM_CHAR, ByVal Asc("("), ByVal &HA0001
    PostMessage HwndJulia, WM_KEYUP, VK_SHIFT, &HC0360001
    PostMessage HwndJulia, WM_KEYUP, VK_9, &HC00A0001

    'Send )
    PostMessage HwndJulia, WM_KEYDOWN, VK_SHIFT, &H360001
    PostMessage HwndJulia, WM_KEYDOWN, VK_0, &HB0001
    PostMessage HwndJulia, WM_CHAR, ByVal Asc(")"), ByVal &HB0001
    PostMessage HwndJulia, WM_KEYUP, VK_SHIFT, &HC0360001
    PostMessage HwndJulia, WM_KEYUP, VK_0, &HC00B0001

    'Send Enter
    PostMessage HwndJulia, WM_KEYDOWN, VK_RETURN, &HC00B0001
    PostMessage HwndJulia, WM_CHAR, ByVal Asc(vbLf), ByVal &H1C0001
    PostMessage HwndJulia, WM_KEYUP, VK_RETURN, &HC01C0001

    Exit Sub
ErrHandler:
    Throw "#SendMessageToJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


Function GetHandleFromPartialCaption(ByRef lwnd As LongPtr, ByVal sCaption As String) As Boolean

          Dim lhWndP As LongPtr
          Dim sStr As String
1         GetHandleFromPartialCaption = False
2         lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
3         Do While lhWndP <> 0
4             sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
5             GetWindowText lhWndP, sStr, Len(sStr)
6             sStr = Left$(sStr, Len(sStr) - 1)
7             If InStr(1, sStr, sCaption) > 0 Then
8                 GetHandleFromPartialCaption = True
9                 lwnd = lhWndP
10                Exit Do
11            End If
12            lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
13        Loop

End Function

