Attribute VB_Name = "ColorPickerDemo"
Option Explicit

Sub DemoNoArgs()
    ColorPicker
    DemoShowColorMessage ColorPickerUtils.SelectedColor
End Sub

Sub DemoOneLong()
    ColorPicker 1234
    DemoShowColorMessage ColorPickerUtils.SelectedColor
End Sub

Sub DemoThreeLongs()
    ColorPicker 128, 255, 128
    DemoShowColorMessage ColorPickerUtils.SelectedColor
End Sub

Sub DemoShowColorMessage(ByVal color As Long)
    Dim myColor As ColorPickerUtils.PickColor
    myColor = ColorPickerUtils.GetRGBFromLong(color)
    
    MsgBox "You chose a color: " & vbCrLf & _
            "Long value = " & color & vbCrLf & _
            "RGB value = (" & myColor.red & _
            ", " & myColor.green & ", " & myColor.blue & ")"
            
End Sub
