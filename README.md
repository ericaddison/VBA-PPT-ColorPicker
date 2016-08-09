This is a simple color picker for PowerPoint VBA. I couldn't find a built-in color chooser for PowerPoint VBA, even though there IS one for Excel VBA (frustrating!), so I made this simple one that provides RGB editing with sliders, theme colors, and some standard colors.

It is fairly straight forward to use. Functions for launching the color chooser are in the ColorPicker module. 

To launch the color chooser without a default color, you just need to call: `ColorPicker`

To set a default color with a single `Long` value, call: `ColorPicker 1234`

To set a default color with red/green/blue values, call: `ColorPicker 100, 100, 100`
