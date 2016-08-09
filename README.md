This is a simple color picker for PowerPoint VBA. I couldn't find a built-in color chooser for PowerPoint VBA, even though there IS one for Excel VBA (frustrating!), so I made this simple one that provides RGB editing with sliders, theme colors, and some standard colors.


### Launching the color picker
It is fairly straight forward to use. Functions for launching the color chooser are in the `ColorPickerUtils` module. 

To launch the color chooser without a default color, you just need to call: `ColorPicker`

To set a default color with a single `Long` value, call: `ColorPicker 1234`

To set a default color with red/green/blue values, call: `ColorPicker 100, 100, 100`

### Getting the color
After the color picker is closed, you can retrieve the selected color with `SelectedColor` from the `ColorPickerUtils` module. The color is returned as a `Long` value, which is how VBA prefers to pass colors around.

### Files
This is a list of the files, in order of relevance:
* ColorPickerForm.frm: The `UserForm` file (along with the .frx file) for the ColorPicker dialog box.
* ColorPickerUtils.bas: The `Module` file with functions to launch the dialog, retrieve the selected color, and convert colors to RGB.
* ColorPickerDemo.bas: A small collection of examples for launching the ColorPicker, with and without initial colors.
* SlideUtils.bas: A couple of simple routines for working with slides.
* ColorPickerForm.frx: The binary file that accomanies the .frm form file.
