Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
    End With
    Range("I6").Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = -0.249977111117893
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    With Selection.Font
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = -0.249977111117893
    End With
End Sub
