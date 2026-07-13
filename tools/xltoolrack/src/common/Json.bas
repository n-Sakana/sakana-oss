Attribute VB_Name = "Json"
Option Explicit
Option Private Module

Public Function JsonQuote(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, Chr$(34), "\" & Chr$(34))
    value = Replace(value, vbCr, "\r")
    value = Replace(value, vbLf, "\n")
    JsonQuote = Chr$(34) & value & Chr$(34)
End Function

