Attribute VB_Name = "ColorPickerDemo"
Sub DemoNoArgs()
    ColorPicker
    MsgBox "You chose a color with Long value: " & ColorPickerUtils.SelectedColor
End Sub

Sub DemoOneLong()
    ColorPicker 1234
    MsgBox "You chose a color with Long value: " & ColorPickerUtils.SelectedColor
End Sub

Sub DemoThreeLongs()
    ColorPicker 128, 255, 128
    MsgBox "You chose a color with Long value: " & ColorPickerUtils.SelectedColor
End Sub
