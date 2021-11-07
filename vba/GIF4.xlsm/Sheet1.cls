VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Static inMethod As Boolean
    If inMethod Then Exit Sub
    Dim Formula As String

    If Target.Areas.Count = 1 Then
        If Target.Cells(1, 1).HasFormula Then
            If Target.Row > 1 Then
                If IsEmpty(Target.Cells(0, 1)) Then
                    inMethod = True
                    With Target.Cells(0, 1)
                        Formula = "=""Formula at " & Replace(.Address, "$", "") & " is "" & " & "FORMULATEXT(R[1]C)"
                        With .Font
                            .Italic = True
                            .ThemeColor = xlThemeColorAccent2
                            .TintAndShade = -0.249977111117893
                        End With
                        .Formula2R1C1 = Formula
                    End With
                    inMethod = False
                End If
            End If
        End If
    End If


End Sub
