VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorPickerForm 
   Caption         =   "Color Picker"
   ClientHeight    =   4620
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4788
   OleObjectBlob   =   "ColorPickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorPickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'*************************************************************
' Attributes
Private SelectedColor As ColorPickerUtils.PickColor
Private MyStandardColors As New Collection

'*************************************************************
' Initilialize

Public Sub resetForm()
    SelectedColor.red = 0
    SelectedColor.green = 0
    SelectedColor.blue = 0
    updateColor
    setStandardColors
    setThemeColors
End Sub


'*************************************************************
' Public functions

Public Function GetSelectedColor() As Long
    If SelectedColor.red = -1 Then
        GetSelectedColor = -1
    Else
        GetSelectedColor = RGB(SelectedColor.red, SelectedColor.green, SelectedColor.blue)
    End If
End Function

Public Sub SetSelectedColor(ByVal color As Long)
    SelectedColor = GetRGBFromLong(color)
    updateColor
End Sub


'*************************************************************
' TextBox functions

Private Sub RedBox_Change()
    RedBox.text = setColor(RedBox.text, SelectedColor.red)
    updateColor
End Sub

Private Sub GreenBox_Change()
    GreenBox.text = setColor(GreenBox.text, SelectedColor.green)
    updateColor
End Sub

Private Sub BlueBox_Change()
    BlueBox.text = setColor(BlueBox.text, SelectedColor.blue)
    updateColor
End Sub

Private Sub LongBox_Change()
    On Error Resume Next
        SelectedColor = GetRGBFromLong(LongBox.text)
        updateColor
    On Error GoTo 0
End Sub



'*************************************************************
' Scrollbar functions

Private Sub RedBar_Change()
    SelectedColor.red = RedBar.Value
    updateColor
End Sub

Private Sub GreenBar_Change()
    SelectedColor.green = GreenBar.Value
    updateColor
End Sub

Private Sub BlueBar_Change()
    SelectedColor.blue = BlueBar.Value
    updateColor
End Sub

Private Sub RedBar_Scroll()
    SelectedColor.red = RedBar.Value
    updateColor
End Sub

Private Sub GreenBar_Scroll()
    SelectedColor.green = GreenBar.Value
    updateColor
End Sub

Private Sub BlueBar_Scroll()
    SelectedColor.blue = BlueBar.Value
    updateColor
End Sub


'*************************************************************
' Button functions

Private Sub OKButton_Click()
    ColorPickerForm.Hide
End Sub

Private Sub CancelButton_Click()
    SelectedColor.red = -1
    SelectedColor.blue = -1
    SelectedColor.green = -1
    ColorPickerForm.Hide
End Sub


'*************************************************************
' Helper functions

' set the color label background color
Private Sub updateColor()
    With SelectedColor
        ColorLabel.BackColor = RGB(.red, .green, .blue)
        LongBox.Value = RGB(.red, .green, .blue)
        RedBox.Value = .red
        RedBar.Value = .red
        GreenBox.Value = .green
        GreenBar.Value = .green
        BlueBox.Value = .blue
        BlueBar.Value = .blue
    End With
End Sub


' set the color to the value parsed from text, with limits of
' 0-255
Private Function setColor(ByRef text As String, ByRef color As Integer) As Integer
    On Error Resume Next
        If text = "" Then
            color = 0
        Else
            color = CInt(text)
            If color < 0 Then
                color = 0
            ElseIf color > 255 Then
                color = 255
            End If
        End If
    On Error GoTo 0
    setColor = color
End Function


'*************************************************************
' Color Array Functions

' set the theme color boxes
Private Sub setThemeColors()
    
    ' Create file path string
    Dim curSlide As Slide
    Set curSlide = SlideUtils.CurrentSlide("")
    If curSlide Is Nothing Then
        ThemeColor1.BackColor = MyStandardColors(1)
        ThemeColor2.BackColor = MyStandardColors(2)
        ThemeColor3.BackColor = MyStandardColors(3)
        ThemeColor4.BackColor = MyStandardColors(4)
        ThemeColor5.BackColor = MyStandardColors(5)
        ThemeColor6.BackColor = MyStandardColors(6)
        ThemeColor7.BackColor = MyStandardColors(7)
        ThemeColor8.BackColor = MyStandardColors(8)
        ThemeColor9.BackColor = MyStandardColors(9)
        ThemeColor10.BackColor = MyStandardColors(10)
        ThemeColor11.BackColor = MyStandardColors(11)
        ThemeColor12.BackColor = MyStandardColors(12)
    Else
        With ActiveWindow.View.Slide.ThemeColorScheme
            ThemeColor1.BackColor = .colors(1)
            ThemeColor2.BackColor = .colors(2)
            ThemeColor3.BackColor = .colors(3)
            ThemeColor4.BackColor = .colors(4)
            ThemeColor5.BackColor = .colors(5)
            ThemeColor6.BackColor = .colors(6)
            ThemeColor7.BackColor = .colors(7)
            ThemeColor8.BackColor = .colors(8)
            ThemeColor9.BackColor = .colors(9)
            ThemeColor10.BackColor = .colors(10)
            ThemeColor11.BackColor = .colors(11)
            ThemeColor12.BackColor = .colors(12)
        End With
    End If
    


End Sub

' set the theme color boxes
Private Sub setStandardColors()
    
        MyStandardColors.Add RGB(128, 0, 0)
        MyStandardColors.Add RGB(255, 0, 0)
        MyStandardColors.Add RGB(255, 128, 0)
        MyStandardColors.Add RGB(255, 255, 0)
        MyStandardColors.Add RGB(0, 128, 0)
        MyStandardColors.Add RGB(0, 255, 0)
        MyStandardColors.Add RGB(0, 0, 128)
        MyStandardColors.Add RGB(0, 0, 255)
        MyStandardColors.Add RGB(0, 255, 255)
        MyStandardColors.Add RGB(255, 0, 255)
        MyStandardColors.Add RGB(100, 100, 100)
        MyStandardColors.Add RGB(200, 200, 200)
    
        StandardColor1.BackColor = MyStandardColors(1)
        StandardColor2.BackColor = MyStandardColors(2)
        StandardColor3.BackColor = MyStandardColors(3)
        StandardColor4.BackColor = MyStandardColors(4)
        StandardColor5.BackColor = MyStandardColors(5)
        StandardColor6.BackColor = MyStandardColors(6)
        StandardColor7.BackColor = MyStandardColors(7)
        StandardColor8.BackColor = MyStandardColors(8)
        StandardColor9.BackColor = MyStandardColors(9)
        StandardColor10.BackColor = MyStandardColors(10)
        StandardColor11.BackColor = MyStandardColors(11)
        StandardColor12.BackColor = MyStandardColors(12)
End Sub

Private Sub setColorFromTheme(ByVal ind As MsoThemeColorSchemeIndex)
    Dim curSlide As Slide
    Set curSlide = SlideUtils.CurrentSlide("")
    If curSlide Is Nothing Then
        ind = ((ind - 1) Mod MyStandardColors.Count) + 1
        SelectedColor = GetRGBFromLong(MyStandardColors(ind))
    Else
        SelectedColor = GetRGBFromLong(curSlide.ThemeColorScheme.colors(ind))
    End If
    updateColor
End Sub

Private Sub setColorFromStandard(ByVal ind As Integer)
    SelectedColor = GetRGBFromLong(MyStandardColors(ind))
    updateColor
End Sub


Private Sub ThemeColor1_Click()
    setColorFromTheme 1
End Sub

Private Sub ThemeColor2_Click()
    setColorFromTheme 2
End Sub

Private Sub ThemeColor3_Click()
    setColorFromTheme 3
End Sub

Private Sub ThemeColor4_Click()
    setColorFromTheme 4
End Sub

Private Sub ThemeColor5_Click()
    setColorFromTheme 5
End Sub

Private Sub ThemeColor6_Click()
    setColorFromTheme 6
End Sub

Private Sub ThemeColor7_Click()
    setColorFromTheme 7
End Sub

Private Sub ThemeColor8_Click()
    setColorFromTheme 8
End Sub

Private Sub ThemeColor9_Click()
    setColorFromTheme 9
End Sub

Private Sub ThemeColor10_Click()
    setColorFromTheme 10
End Sub

Private Sub ThemeColor11_Click()
    setColorFromTheme 11
End Sub

Private Sub ThemeColor12_Click()
    setColorFromTheme 12
End Sub

Private Sub StandardColor1_Click()
    setColorFromStandard 1
End Sub

Private Sub StandardColor2_Click()
    setColorFromStandard 2
End Sub

Private Sub StandardColor3_Click()
    setColorFromStandard 3
End Sub

Private Sub StandardColor4_Click()
    setColorFromStandard 4
End Sub

Private Sub StandardColor5_Click()
    setColorFromStandard 5
End Sub

Private Sub StandardColor6_Click()
    setColorFromStandard 6
End Sub

Private Sub StandardColor7_Click()
    setColorFromStandard 7
End Sub

Private Sub StandardColor8_Click()
    setColorFromStandard 8
End Sub

Private Sub StandardColor9_Click()
    setColorFromStandard 9
End Sub

Private Sub StandardColor10_Click()
    setColorFromStandard 10
End Sub

Private Sub StandardColor11_Click()
    setColorFromStandard 11
End Sub

Private Sub StandardColor12_Click()
    setColorFromStandard 12
End Sub

