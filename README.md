<div align="center">

## Boxgradient module


</div>

### Description

A very neat effect to color your forms or pictureboxes. This code is very

small but very usefull. A MUST HAVE for anyone who likes different forms...
 
### More Info
 
The routine must be stored in a module.

This effect can be done in a form or a picturebox. If you call it

in the load-event, remember to set the autoredraw property to true.

Syntax: Call BoxGradient(Object,r,g,b,rstep,gstep,bstep,direc)

where: object = the form or picturebox

r,g,b = the starting-colors of the gradient

rstep,gstep,bstep: the amount of increasing the colors

direc: true or false

For example:

Call BoxGradient(Form1,128, 64, 0, 1, 2, 2, False)

stephan.swertvaegher@planetinternet.be


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[stephane swertvaegher](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephane-swertvaegher.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephane-swertvaegher-boxgradient-module__1-3471/archive/master.zip)





### Source Code

```
Option Explicit
Public Sub BoxGradient(OBJ As Object, R%, G%, B%, RStep%, GStep%, BStep%, Direc As Boolean)
Dim s%, xpos%, ypos%
OBJ.ScaleMode = 3 'pixel
If Direc = True Then
RStep% = -RStep%
GStep% = -GStep%
BStep% = -BStep%
End If
DoBox:
s% = s% + 1
If xpos% < Int(OBJ.ScaleWidth / 2) Then xpos% = s%
If ypos% < Int(OBJ.ScaleHeight / 2) Then ypos% = s%
OBJ.Line (xpos%, ypos%)-(OBJ.ScaleWidth - xpos%, OBJ.ScaleHeight - ypos%), RGB(R%, G%, B%), B
R% = R% - RStep%
If R% < 0 Then R% = 0
If R% > 255 Then R% = 255
G% = G% - GStep%
If G% < 0 Then G% = 0
If G% > 255 Then G% = 255
B% = B% - BStep%
If B% < 0 Then B% = 0
If B% > 255 Then B% = 255
If xpos% = Int(OBJ.ScaleWidth / 2) And ypos% = Int(OBJ.ScaleHeight / 2) Then
Exit Sub
End If
GoTo DoBox
End Sub
```

