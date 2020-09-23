<div align="center">

## Resizer Reupload


</div>

### Description

This class resizes ALL the controls on a form in just two lines of code. No API, no OCX, just a simple class, very easy to use. It tests if a control is resizable, and when ok, the control (and optionally the font) will be resized everytime the form is resized. I have noticed that the clsresizer class has to be compiled if a dll is not found...

usage :

make a new instance of the class

dim r as new resizer

in form_load :

r.ResizeFonts = True (optional, this line starts the font resizing)

set r.frmToSize = me 'Required

r.AddItemNotToResize "command1" (optional : do not resize command1)

...

end sub

in form_resize :

r.adjustSize 'resizes the controls (and the fonts if selected)

that's all!

optional : r.ResetForm(when dragging / dropping ) ...

-> Get the new sizes (does not resize a control (to return to original sizes use r.adjustsize instead)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-06-13 08:44:24
**By**             |[Sven Maes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sven-maes.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Resizer\_Re938566132002\.zip](https://github.com/Planet-Source-Code/sven-maes-resizer-reupload__1-35783/archive/master.zip)








