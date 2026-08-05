Attribute VB_Name = "Module1"
Option Explicit

Sub Testsasdc()

Dim Foo

Foo = JuliaEvalVBA("mda(9)")

Debug.Print NumDimensions(Foo)

End Sub
