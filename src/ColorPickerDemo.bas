Attribute VB_Name = "ColorPickerDemo"
Option Explicit

Sub DemoNoArgs()
    Dim myColor As Long
    myColor = ColorPicker
    DemoShowColorMessage myColor
End Sub

Sub DemoOneLong()
    DemoShowColorMessage ColorPicker(1234)
End Sub

Sub DemoThreeLongs()
    ColorPicker 128, 255, 128
    DemoShowColorMessage ColorPickerUtils.SelectedColor
End Sub

Sub DemoClickShape(shp As Shape)
    Dim newColor As Long
    newColor = ColorPicker(shp.Fill.ForeColor.RGB)
    shp.Fill.ForeColor.RGB = newColor
End Sub


Sub DemoShowColorMessage(ByVal color As Long)
    Dim myColor As ColorPickerUtils.PickColor
    myColor = ColorPickerUtils.GetRGBFromLong(color)
    
    MsgBox "You chose a color: " & vbCrLf & _
            "Long value = " & color & vbCrLf & _
            "RGB value = (" & myColor.red & _
            ", " & myColor.green & ", " & myColor.blue & ")"
            
End Sub

