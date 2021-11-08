Attribute VB_Name = "modAdHoc"
Option Explicit

Sub test()

    Dim RefersTo

    On Error GoTo ErrHandler
    ThrowIfError JuliaLaunch

    RefersTo = Array(1, 2, Array(3, 4))

    ThrowIfError JuliaSetVar("foo", RefersTo)

    Exit Sub
ErrHandler:
    MsgBox "#test (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
