This is a simple color picker for PowerPoint VBA. I couldn't find a built-in color chooser for PowerPoint VBA, even though there IS one for Excel VBA (frustrating!), so I made this simple one that provides RGB editing with sliders, theme colors, and some standard colors.

<p align="center">
  <img src="https://github.com/ericaddison/VBA-PPT-ColorPicker/blob/gh-pages/img/screenShot.png?raw=true" alt="VBA-PPT-ColorPicker ScreenShot"/>
</p>

### Launching the color picker
It is fairly straight forward to use. Functions for launching the color chooser are in the `ColorPickerUtils` module:
```
ColorPicker                ' Launch a ColorPicker with no initial color
ColorPicker 1234           ' Launch a ColorPicker with initial color 1234
ColorPicker 100, 100, 100  ' Launch a ColorPicker with initial color RGB(100, 100, 100)
```

### Getting the color
There are two ways to retrieve the selected color. `ColorPicker` is a function that returns a the color as a `Long` value, which is how VBA prefers to pass colors around. Additionally, there is a `SelectedColor` function in `ColorPickerUtils` that returns the selected color as well:

``` 
Dim myColor As Long
myColor = ColorPicker                     ' assign selected color directly from call to ColorPicker
myColor = ColorPickerUtils.SelectedColor  ' assign from helper function
```

`SelectedColor` will return as `-1` if the selection was cancelled.

### Alternative
It is possible to launch a Microsoft system color dialog from `comdlg32.dll` which provides a very nice color chooser interface. Unfortunately, I was not able to get this dialog to set an initial color, or to position properly due to the inability to easily access the window handle for a UserForm or Window in VBA for PowerPoint. If you want to give it a try, though, here is a link to an old Microsoft article with code: [How To Use Color Dialog from COMDLG32.DLL](https://support.microsoft.com/en-us/kb/153929). 

### Files
This is a list of the files, in order of relevance:
* ColorPickerForm.frm: The `UserForm` file (along with the .frx file) for the ColorPicker dialog box.
* ColorPickerUtils.bas: The `Module` file with functions to launch the dialog, retrieve the selected color, and convert colors to RGB.
* ColorPickerDemo.bas: A small collection of examples for launching the ColorPicker, with and without initial colors.
* SlideUtils.bas: A couple of simple routines for working with slides.
* ColorPickerForm.frx: The binary file that accomanies the .frm form file.
