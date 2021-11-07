Attribute VB_Name = "modClipboard"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modClipboard
' Author    : Philip Swannell
' Date      : 28-Mar-2016, 21-Nov-2017
' Purpose   : Copy text to clipboard, Try two approaches: via a "DataObject" and via Windows API
'             References:
'             http://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
'             21 Nov 2017 Was finding problems with 64-bit implementation on the page above, fixed on Stack Overflow :-)
'             https://stackoverflow.com/questions/18668928/excel-2013-64-bit-vba-clipboard-api-doesnt-work
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CopyStringToClipboard
' Author     : Philip Swannell
' Date       : 08-Feb-2018
' Purpose    : Copies a string to the Windows clipboard. Previously had a way of doing this using DataObject.PutInClipboard, but that proved unreliable
' Parameters :
'  MyString:  The string to put on the clipboard
' -----------------------------------------------------------------------------------------------------------------------
Sub CopyStringToClipboard(MyString As String)

          Dim ErrorMessage As String
          Dim hClipMemory As LongPtr
          Dim hGlobalMemory As LongPtr
          Dim lpGlobalMemory As LongPtr
          Dim x As Long
          Const GHND = &H42
          Const CF_TEXT = 1

          'Allocate moveable global memory
1         On Error GoTo ErrHandler
2         hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

          'Lock the block to get a far pointer to this memory.
3         lpGlobalMemory = GlobalLock(hGlobalMemory)

          'Copy the string to this global memory.
4         lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

          'Unlock the memory.
5         If GlobalUnlock(hGlobalMemory) <> 0 Then
6             ErrorMessage = "Could not unlock memory location. Copy aborted."
7             GoTo OutOfHere2
8         End If

          'Open the Clipboard to copy data to.
9         If OpenClipboard(0&) = 0 Then
10            Throw "Could not open the Clipboard. Copy aborted."
11        End If

          'Clear the Clipboard.
12        x = EmptyClipboard()

          'Copy the data to the Clipboard.
13        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
14        If CloseClipboard() = 0 Then
15            ErrorMessage = "Could not close Clipboard."
16        End If

17        If ErrorMessage <> vbNullString Then Throw ErrorMessage
18        Exit Sub
ErrHandler:
19        Throw "#CopyStringToClipboard (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
