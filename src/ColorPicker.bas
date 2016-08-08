Attribute VB_Name = "ColorPicker"

' Launch a color picker form in one of three ways:
' 1) no arguments: initial color is black
' 2) 1 Long argument: intial color set with Long color value
' 3) 3 Long arguments: initial color set with RGB(red,green,blue)
Public Sub ColorPicker(Optional ByVal red As Long = -1, _
        Optional ByVal green As Long = -1, Optional ByVal blue As Long = -1)
    
    If Not red = -1 Then
        Load ColorPickerForm
        If (green = -1 And blue = -1) Then
            ColorPickerForm.SetSelectedColor red
        Else
            ColorPickerForm.SetSelectedColor RGB(red, green, blue)
        End If
    End If
    ColorPickerForm.Show
End Sub


Sub DemoNoArgs()
    ColorPicker
End Sub

Sub DemoOneLong()
    ColorPicker 1234
End Sub

Sub DemoThreeLongs()
    ColorPicker 128, 255, 128
End Sub

