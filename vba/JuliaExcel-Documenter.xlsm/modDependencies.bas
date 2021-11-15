Attribute VB_Name = "modDependencies"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsInCollection
' Purpose   : Tests for membership of any collection. Can be used in place of multiple
'             functions BookHasSheet, SheetHasName etc. Works irrespective of whether
'             the collection contains objects or primitives.
' -----------------------------------------------------------------------------------------------------------------------
Function IsInCollection(oColn As Object, Key As String) As Boolean
1         On Error GoTo ErrHandler
2         VarType (oColn(Key))
3         IsInCollection = True
4         Exit Function
ErrHandler:
5     End Function

Function RawFileContents(FileName As String)
          Dim F As Scripting.File
          Dim FSO As New FileSystemObject
          Dim T As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set F = FSO.GetFile(FileName)
3         Set T = F.OpenAsTextStream()
4         RawFileContents = T.ReadAll
5         T.Close

6         Exit Function
ErrHandler:
           Throw "#RawFileContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateStacker
' Purpose   : So that we can create clsStacker objects from other workbooks...
' -----------------------------------------------------------------------------------------------------------------------
Function CreateStacker() As clsStacker
1         On Error GoTo ErrHandler
2         Set CreateStacker = New clsStacker
3         Exit Function
ErrHandler:
4         Throw "#CreateStacker (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateMissing
' Purpose   : Returns a variant of type Missing
'---------------------------------------------------------------------------------------
Private Function CreateMissing()
1         CreateMissing = CM2()
End Function
Private Function CM2(Optional OptionalArg As Variant)
1         CM2 = OptionalArg
End Function

Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim y As Long
1         If Not IsArray(x) Then
2             NumDimensions = 0
3             Exit Function
4         Else
5             On Error GoTo ExitPoint
6             i = 1
7             Do While True
8                 y = LBound(x, i)
9                 i = i + 1
10            Loop
11        End If
ExitPoint:
12        NumDimensions = i - 1
End Function
'---------------------------------------------------------------------------------------
' Procedure : sNRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
'---------------------------------------------------------------------------------------
Public Function sNRows(Optional TheArray) As Long
1         If TypeName(TheArray) = "Range" Then
2             sNRows = TheArray.Rows.Count
3         ElseIf IsMissing(TheArray) Then
4             sNRows = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             sNRows = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     sNRows = 1
10                Case Else
11                    sNRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
12            End Select
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
' -----------------------------------------------------------------------------------------------------------------------
Function sNCols(Optional TheArray) As Long
1         If TypeName(TheArray) = "Range" Then
2             sNCols = TheArray.Columns.Count
3         ElseIf IsMissing(TheArray) Then
4             sNCols = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             sNCols = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     sNCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
10                Case Else
11                    sNCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
12            End Select
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Throw(ByVal ErrorString As String)
1         Err.Raise vbObjectError + 1, , ErrorString
End Sub

Public Function StringBetweenStrings(TheString, LeftString, RightString, Optional IncludeLeftString As Boolean, Optional IncludeRightString As Boolean)
          Dim MatchPoint1 As Long        ' the position of the first character to return
          Dim MatchPoint2 As Long        ' the position of the last character to return
          Dim FoundLeft As Boolean
          Dim FoundRight As Boolean
          
1         On Error GoTo ErrHandler
          
2         If VarType(TheString) <> vbString Or VarType(LeftString) <> vbString Or VarType(RightString) <> vbString Then Throw "Inputs must be strings"
3         If LeftString = vbNullString Then
4             MatchPoint1 = 0
5         Else
6             MatchPoint1 = InStr(1, TheString, LeftString, vbTextCompare)
7         End If

8         If MatchPoint1 = 0 Then
9             FoundLeft = False
10            MatchPoint1 = 1
11        Else
12            FoundLeft = True
13        End If

14        If RightString = vbNullString Then
15            MatchPoint2 = 0
16        ElseIf FoundLeft Then
17            MatchPoint2 = InStr(MatchPoint1 + Len(LeftString), TheString, RightString, vbTextCompare)
18        Else
19            MatchPoint2 = InStr(1, TheString, RightString, vbTextCompare)
20        End If

21        If MatchPoint2 = 0 Then
22            FoundRight = False
23            MatchPoint2 = Len(TheString)
24        Else
25            FoundRight = True
26            MatchPoint2 = MatchPoint2 - 1
27        End If

28        If Not IncludeLeftString Then
29            If FoundLeft Then
30                MatchPoint1 = MatchPoint1 + Len(LeftString)
31            End If
32        End If

33        If IncludeRightString Then
34            If FoundRight Then
35                MatchPoint2 = MatchPoint2 + Len(RightString)
36            End If
37        End If

38        StringBetweenStrings = Mid$(TheString, MatchPoint1, MatchPoint2 - MatchPoint1 + 1)

39        Exit Function
ErrHandler:
40        StringBetweenStrings = "#StringBetweenStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
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
