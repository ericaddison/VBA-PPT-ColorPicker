Attribute VB_Name = "ColorPickerUtils"
Option Explicit

Public Type PickColor
    red As Integer
    green As Integer
    blue As Integer
End Type

Private pColor As Long


' Launch a color picker form in one of three ways:
' 1) no arguments: initial color is black
' 2) 1 Long argument: intial color set with Long color value
' 3) 3 Long arguments: initial color set with RGB(red,green,blue)
Public Function ColorPicker(Optional ByVal red As Long = -1, _
        Optional ByVal green As Long = -1, Optional ByVal blue As Long = -1) As Long
    Load ColorPickerForm
    ColorPickerForm.resetForm
    If Not red = -1 Then
        If (green = -1 And blue = -1) Then
            ColorPickerForm.SetSelectedColor red
        Else
            ColorPickerForm.SetSelectedColor RGB(red, green, blue)
        End If
    End If
    ColorPickerForm.Show
    ColorPicker = ColorPickerForm.GetSelectedColor
    pColor = ColorPickerForm.GetSelectedColor
    Unload ColorPickerForm
End Function

Public Function SelectedColor() As Long
    SelectedColor = pColor
End Function

' get separate R-G-B values from a color stored as a Long
Public Function GetRGBFromLong(ByVal color As Long) As PickColor
    Dim newColor As PickColor
    newColor.red = color Mod 256
    newColor.green = color \ 256 Mod 256
    newColor.blue = color \ (65536) Mod 256
    GetRGBFromLong = newColor
End Function






