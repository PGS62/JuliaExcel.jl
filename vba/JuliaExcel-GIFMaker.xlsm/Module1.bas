Attribute VB_Name = "Module1"
Option Explicit

Sub MakeGIF()
          Dim Keys As Variant
          Dim Key As String
          Dim i As Long, j As Long
          Dim t1 As Double
          Dim DelayAfter As Double
          Dim DelayBetween As Double
          Dim waitTime As Double

1         On Error GoTo ErrHandler
2         AppActivate "Julia 1.6.3"
3         ActiveSheet.Calculate
4         Keys = sExpandDown(ActiveSheet.Range("keys")).Value

5         For i = 1 To sNRows(Keys)
6             DelayAfter = ActiveSheet.Range("DelayAfter").Offset(i - 1).Value
7             DelayBetween = ActiveSheet.Range("DelayBetween").Offset(i - 1).Value
8             Key = Keys(i, 1)

9             If DelayBetween <= 0 Or InStr(Key, "{") > 0 Then
10                Application.SendKeys Keys(i, 1)
11            Else
12                For j = 1 To Len(Key)
13                    Application.SendKeys Mid$(Key, j, 1)
14                    t1 = sElapsedTime()
15                    While sElapsedTime() < t1 + DelayBetween
16                        DoEvents
17                    Wend
18                Next j
19            End If
20            If DelayAfter > 0 Then
21                t1 = sElapsedTime()


22                Do
23                    waitTime = DelayAfter - sElapsedTime() + t1
24                    If waitTime <= 0 Then
25                        Application.StatusBar = False
26                        DoEvents
27                        Exit Do
28                    End If

29                    Application.StatusBar = Format(waitTime, "#0")
30                    DoEvents
31                Loop
32            End If
33        Next i

34        Exit Sub
ErrHandler:
35        MsgBox "#MakeGIF (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


