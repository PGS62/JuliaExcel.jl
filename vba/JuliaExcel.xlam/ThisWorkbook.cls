VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AddinInstall()
  '  Debug.Print "Calling RegisterAll from Workbook_AddinInstall"
    RegisterAll
End Sub

Private Sub Workbook_Open()
  '  Debug.Print "Calling RegisterAll from Workbook_Open"
    RegisterAll
End Sub



