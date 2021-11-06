Attribute VB_Name = "modDependencies"
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------------------------
' Procedure : VStack
' Purpose   : Places arrays on top of one another. If the arrays are of unequal width then they will be
'             padded to the right with #NA! values.
' Arguments
' ArraysToStack:
'---------------------------------------------------------------------------------------------------------
Public Function VStack(ParamArray ArraysToStack())
    Dim AllC As Long
    Dim AllR As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim R As Long
    Dim R0 As Long
    Dim ReturnArray()
    On Error GoTo ErrHandler

    Static NA As Variant
    If IsMissing(ArraysToStack) Then
        VStack = CreateMissing()
    Else
        If IsEmpty(NA) Then NA = CVErr(xlErrNA)

        For i = LBound(ArraysToStack) To UBound(ArraysToStack)
            If TypeName(ArraysToStack(i)) = "Range" Then ArraysToStack(i) = ArraysToStack(i).value
            If IsMissing(ArraysToStack(i)) Then
                R = 0: c = 0
            Else
                Select Case NumDimensions(ArraysToStack(i))
                    Case 0
                        R = 1: c = 1
                    Case 1
                        R = 1
                        c = UBound(ArraysToStack(i)) - LBound(ArraysToStack(i)) + 1
                    Case 2
                        R = UBound(ArraysToStack(i), 1) - LBound(ArraysToStack(i), 1) + 1
                        c = UBound(ArraysToStack(i), 2) - LBound(ArraysToStack(i), 2) + 1
                End Select
            End If
            If c > AllC Then AllC = c
            AllR = AllR + R
        Next i

        If AllR = 0 Then
            VStack = CreateMissing
            Exit Function
        End If

        ReDim ReturnArray(1 To AllR, 1 To AllC)

        R0 = 1
        For i = LBound(ArraysToStack) To UBound(ArraysToStack)
            If Not IsMissing(ArraysToStack(i)) Then
                Select Case NumDimensions(ArraysToStack(i))
                    Case 0
                        R = 1: c = 1
                        ReturnArray(R0, 1) = ArraysToStack(i)
                    Case 1
                        R = 1
                        c = UBound(ArraysToStack(i)) - LBound(ArraysToStack(i)) + 1
                        For j = 1 To c
                            ReturnArray(R0, j) = ArraysToStack(i)(j + LBound(ArraysToStack(i)) - 1)
                        Next j
                    Case 2
                        R = UBound(ArraysToStack(i), 1) - LBound(ArraysToStack(i), 1) + 1
                        c = UBound(ArraysToStack(i), 2) - LBound(ArraysToStack(i), 2) + 1

                        For j = 1 To R
                            For k = 1 To c
                                ReturnArray(R0 + j - 1, k) = ArraysToStack(i)(j + LBound(ArraysToStack(i), 1) - 1, k + LBound(ArraysToStack(i), 2) - 1)
                            Next k
                        Next j

                End Select
                If c < AllC Then
                    For j = 1 To R
                        For k = c + 1 To AllC
                            ReturnArray(R0 + j - 1, k) = NA
                        Next k
                    Next j
                End If
                R0 = R0 + R
            End If
        Next i

        VStack = ReturnArray
    End If
    Exit Function
ErrHandler:
    VStack = "#VStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateMissing
' Purpose   : Returns a variant of type Missing
'---------------------------------------------------------------------------------------
Private Function CreateMissing()
    CreateMissing = CM2()
End Function
Private Function CM2(Optional OptionalArg As Variant)
    CM2 = OptionalArg
End Function

Private Function NumDimensions(x As Variant) As Long
    Dim i As Long
    Dim y As Long
    If Not IsArray(x) Then
        NumDimensions = 0
        Exit Function
    Else
        On Error GoTo ExitPoint
        i = 1
        Do While True
            y = LBound(x, i)
            i = i + 1
        Loop
    End If
ExitPoint:
    NumDimensions = i - 1
End Function
'---------------------------------------------------------------------------------------
' Procedure : sNRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
'---------------------------------------------------------------------------------------
Public Function sNRows(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        sNRows = TheArray.Rows.Count
    ElseIf IsMissing(TheArray) Then
        sNRows = 0
    ElseIf VarType(TheArray) < vbArray Then
        sNRows = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                sNRows = 1
            Case Else
                sNRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
        End Select
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub


Public Function StringBetweenStrings(TheString, LeftString, RightString, Optional IncludeLeftString As Boolean, Optional IncludeRightString As Boolean)
    Dim MatchPoint1 As Long        ' the position of the first character to return
    Dim MatchPoint2 As Long        ' the position of the last character to return
    Dim FoundLeft As Boolean
    Dim FoundRight As Boolean
    
    On Error GoTo ErrHandler
    
    If VarType(TheString) <> vbString Or VarType(LeftString) <> vbString Or VarType(RightString) <> vbString Then Throw "Inputs must be strings"
    If LeftString = vbNullString Then
        MatchPoint1 = 0
    Else
        MatchPoint1 = InStr(1, TheString, LeftString, vbTextCompare)
    End If

    If MatchPoint1 = 0 Then
        FoundLeft = False
        MatchPoint1 = 1
    Else
        FoundLeft = True
    End If

    If RightString = vbNullString Then
        MatchPoint2 = 0
    ElseIf FoundLeft Then
        MatchPoint2 = InStr(MatchPoint1 + Len(LeftString), TheString, RightString, vbTextCompare)
    Else
        MatchPoint2 = InStr(1, TheString, RightString, vbTextCompare)
    End If

    If MatchPoint2 = 0 Then
        FoundRight = False
        MatchPoint2 = Len(TheString)
    Else
        FoundRight = True
        MatchPoint2 = MatchPoint2 - 1
    End If

    If Not IncludeLeftString Then
        If FoundLeft Then
            MatchPoint1 = MatchPoint1 + Len(LeftString)
        End If
    End If

    If IncludeRightString Then
        If FoundRight Then
            MatchPoint2 = MatchPoint2 + Len(RightString)
        End If
    End If

    StringBetweenStrings = Mid$(TheString, MatchPoint1, MatchPoint2 - MatchPoint1 + 1)

    Exit Function
ErrHandler:
    StringBetweenStrings = "#StringBetweenStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : sRegExReplace
' Purpose   : Uses regular expressions to make replacement in a set of input strings.
'
'             The function replaces every instance of the regular expression match with the
'             replacement.
' Arguments
' InputString: Input string to be transformed. Can be an array. Non-string elements will be left
'             unchanged.
' RegularExpression: A standard regular expression string.
' Replacement: A replacement template for each match of the regular expression in the input string.
' CaseSensitive: Whether matching should be case-sensitive (TRUE) or not (FALSE).
'
' Notes     : Details of regular expressions are given under sIsRegMatch. The replacement string can be
'             an explicit string, and it can also contain special escape sequences that are
'             replaced by the characters they represent. The options available are:
'
'             Characters Replacement
'             $n        n-th backreference. That is, a copy of the n-th matched group
'             specified with parentheses in the regular expression. n must be an integer
'             value designating a valid backreference, greater than zero, and of two digits
'             at most.
'             $&       A copy of the entire match
'             $`        The prefix, that is, the part of the target sequence that precedes
'             the match.
'             $´        The suffix, that is, the part of the target sequence that follows
'             the match.
'             $$        A single $ character.
'---------------------------------------------------------------------------------------------------------
Function sRegExReplace(InputString As Variant, RegularExpression As String, Replacement As String, Optional CaseSensitive As Boolean)
          Dim i As Long
          Dim j As Long
          Dim Result() As String
          Dim rx As VBScript_RegExp_55.RegExp
1         On Error GoTo ErrHandler

2         If Not RegExSyntaxValid(RegularExpression) Then
3             sRegExReplace = "#Invalid syntax for RegularExpression!"
4             Exit Function
5         End If

6         Set rx = New RegExp

7         With rx
8             .IgnoreCase = Not (CaseSensitive)
9             .Pattern = RegularExpression
10            .Global = True
11        End With

12        If VarType(InputString) = vbString Then
13            sRegExReplace = rx.Replace(InputString, Replacement)
14            GoTo Cleanup
15        ElseIf VarType(InputString) < vbArray Then
16            sRegExReplace = InputString
17            GoTo Cleanup
18        End If
19        If TypeName(InputString) = "Range" Then InputString = InputString.Value2

20        Select Case NumDimensions(InputString)
              Case 2
21                ReDim Result(LBound(InputString, 1) To UBound(InputString, 1), LBound(InputString, 2) To UBound(InputString, 2))
22                For i = LBound(InputString, 1) To UBound(InputString, 1)
23                    For j = LBound(InputString, 2) To UBound(InputString, 2)
24                        If VarType(InputString(i, j)) = vbString Then
25                            Result(i, j) = rx.Replace(InputString(i, j), Replacement)
26                        Else
27                            Result(i, j) = InputString(i, j)
28                        End If
29                    Next j
30                Next i
31            Case 1
32                ReDim Result(LBound(InputString, 1) To UBound(InputString, 1))
33                For i = LBound(InputString, 1) To UBound(InputString, 1)
34                    If VarType(InputString(i)) = vbString Then
35                        Result(i) = rx.Replace(InputString(i), Replacement)
36                    Else
37                        Result(i) = InputString(i)
38                    End If
39                Next i
40            Case Else
41                Throw "InputString must be a String or an array with 1 or 2 dimensions"
42        End Select
43        sRegExReplace = Result

Cleanup:
44        Set rx = Nothing
45        Exit Function
ErrHandler:
46        sRegExReplace = "#sRegExReplace (line " & CStr(Erl) + "): " & Err.Description & "!"
47        Set rx = Nothing
End Function

Function RegExSyntaxValid(RegularExpression As String) As Boolean
          Dim Res As Boolean
          Dim rx As VBScript_RegExp_55.RegExp
1         On Error GoTo ErrHandler
2         Set rx = New RegExp
3         With rx
4             .IgnoreCase = False
5             .Pattern = RegularExpression
6             .Global = False        'Find first match only
7         End With
8         Res = rx.Test("Foo")
9         RegExSyntaxValid = True
10        Exit Function
ErrHandler:
11        RegExSyntaxValid = False
End Function

