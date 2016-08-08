Attribute VB_Name = "SlideUtils"
' Various slide utilities that warrant their own function
Option Explicit
Private Const NO_SLIDE_IN_VIEW_ERROR_CODE As Long = -2147188160

' Insert a new slide in the active presentation one slide ahead
' of the current slide
Function InsertSlide() As Slide
    ' make a new slide
    Dim curSlide As Slide
    Set curSlide = CurrentSlide

    Set InsertSlide = ActivePresentation.Slides.AddSlide(curSlide.SlideIndex + 1, curSlide.CustomLayout)
    ActiveWindow.View.GotoSlide InsertSlide.SlideIndex
End Function


' get the current slide
Function CurrentSlide() As Slide
    Dim curSlideIndex As Integer
    On Error Resume Next
        curSlideIndex = ActiveWindow.View.Slide.SlideIndex
        If Err.Number = NO_SLIDE_IN_VIEW_ERROR_CODE Then
            curSlideIndex = 1
        End If
    On Error GoTo 0
    Set CurrentSlide = ActivePresentation.Slides(curSlideIndex)
End Function
