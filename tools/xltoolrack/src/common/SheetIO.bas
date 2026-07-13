Attribute VB_Name = "SheetIO"
Option Explicit
Option Private Module

Public Sub WriteValues(ByVal target As Range, ByVal values As Variant)
    target.Value2 = values
End Sub

Public Function ReadValues(ByVal source As Range) As Variant
    ReadValues = source.Value2
End Function
