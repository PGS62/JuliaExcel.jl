Attribute VB_Name = "Module1"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sElapsedTime
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Function sElapsedTime() As Double
          Dim a As Currency
          Dim b As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter a
3         QueryPerformanceFrequency b
4         sElapsedTime = a / b

5         Exit Function
ErrHandler:
6         Err.Raise vbObjectError + 1, , "#sElapsedTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function AddCurleys(x As String)
1     Select Case x
      Case "=", "(", ")", "+", "-", "*"
2     AddCurleys = "{" & x & "}"
3     Case Else
4     AddCurleys = x
5     End Select
End Function


Sub MakeGIF()
          Dim Keys As Variant
          Dim Key As String
          Dim i As Long, j As Long
          Dim t1 As Double
          Dim DelayAfter As Double
          Dim DelayBetween As Double
          Dim waitTime As Double
          Dim OtherCommand As String

1         On Error GoTo ErrHandler
2         AppActivate ActiveSheet.Range("WindowToActivate").Value
3         ActiveSheet.Calculate
4         With ActiveSheet.Range("keys")
5             Keys = Range(.Offset(0, 0), .Offset(-1, 0).End(xlDown)).Value
6         End With

7         For i = 1 To UBound(Keys, 1)
8             DelayAfter = ActiveSheet.Range("DelayAfter").Offset(i - 1).Value
9             DelayBetween = ActiveSheet.Range("DelayBetween").Offset(i - 1).Value
10            OtherCommand = ActiveSheet.Range("OtherCommand").Offset(i - 1).Value
11            Key = Keys(i, 1)

12            If InStr(Key, "{") > 0 Then
13                Application.SendKeys Keys(i, 1)
14            Else
15                For j = 1 To Len(Key)
16                    Application.SendKeys AddCurleys(Mid$(Key, j, 1))
17                    If DelayBetween > 0 Then
18                        t1 = sElapsedTime()
19                        While sElapsedTime() < t1 + DelayBetween
20                            DoEvents
21                        Wend
22                    End If
23                Next j
24            End If
25            If DelayAfter > 0 Then
26                t1 = sElapsedTime()
27                Do
28                    waitTime = DelayAfter - sElapsedTime() + t1
29                    If waitTime <= 0 Then
30                        Application.StatusBar = False
31                        DoEvents
32                        Exit Do
33                    End If

34                    Application.StatusBar = Format(waitTime, "#0")
35                    DoEvents
36                Loop
37            End If
38          DoOtherCommand OtherCommand
39        Next i

40        Exit Sub
ErrHandler:
41        MsgBox "#MakeGIF (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub DoOtherCommand(Command As String)
1         On Error GoTo ErrHandler
2         Select Case LCase(Command)

              Case LCase("ReActivateWindow")
                  AppActivate ActiveWindow.Caption
3                 AppActivate ActiveSheet.Range("WindowToActivate").Value
4             Case ""
                  'Do nothing
5             Case Else
6                 Err.Raise vbObjectError + 1, , "Command '" + Command + "' not recognised"
7         End Select

8         Exit Sub
ErrHandler:
9         Err.Raise vbObjectError + 1, , "#DoOtherCommand (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

