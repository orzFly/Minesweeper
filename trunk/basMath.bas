Attribute VB_Name = "basMath"
Option Explicit

Public Function Max(ParamArray Values() As Variant) As Variant
    Dim value As Variant
    If UBound(Values) >= 0 Then Max = Values(0)
    For Each value In Values
        If value > Max Then Max = value
    Next
End Function

Public Function Min(ParamArray Values() As Variant) As Variant
    Dim value As Variant
    If UBound(Values) >= 0 Then Min = Values(0)
    For Each value In Values
        If value < Min Then Min = value
    Next
End Function

Public Function Hex2Dec(ByVal strHex As String) As Long
    Hex2Dec = Int(Val("&H" & strHex))
End Function

Public Function Dec2Hex(ByVal lngDec As Long) As String
    Dec2Hex = Hex(lngDec)
End Function
