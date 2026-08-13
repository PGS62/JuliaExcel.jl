Attribute VB_Name = "modJuliaLiteral"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeJuliaLiteral
' Purpose    : Escape a VBA string so it can be embedded, double-quoted, as a Julia string literal.
'              The sole caller is JuliaLaunch, which uses this to embed the launch command line in a
'              println statement written into the generated Julia startup script.
' -----------------------------------------------------------------------------------------------------------------------
Function MakeJuliaLiteral(x As String) As String
          Dim k As Long
          Dim Res As String

1         Res = x
          'Must do this substitution first
2         If InStr(x, "\") > 0 Then
3             Res = Replace(Res, "\", "\\")
4         End If
          'The conversions in the two loops below are needed to avoid an error: _
          Base.Meta.ParseError("unbalanced bidirectional formatting in string literal") _
          'Julia's "caution" in relation to these characters is a defence against "Trojan Source" attacks.
          'https://github.com/JuliaLang/julia/pull/42918
          'https://trojansource.codes/
5         For k = 8234 To 8238
6             If InStr(x, ChrW(k)) Then
7                 Res = Replace(Res, ChrW(k), "\u" & LCase(Hex(k)))
8             End If
9         Next k
10        For k = 8294 To 8297
11            If InStr(x, ChrW(k)) Then
12                Res = Replace(Res, ChrW(k), "\u" & LCase(Hex(k)))
13            End If
14        Next k
15        If InStr(x, vbCr) > 0 Then
16            Res = Replace(Res, vbCr, "\r")
17        End If
18        If InStr(x, vbLf) > 0 Then
19            Res = Replace(Res, vbLf, "\n")
20        End If
21        If InStr(x, "$") > 0 Then
22            Res = Replace(Res, "$", "\$")
23        End If
24        If InStr(x, """") > 0 Then
25            Res = Replace(Res, """", "\""")
26        End If
27        MakeJuliaLiteral = """" & Res & """"
End Function

