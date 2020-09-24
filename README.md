<div align="center">

## Simplest Translucent \(Semi\-Transparent\) Form


</div>

### Description

This is the simplest way to make a translucent form. I actually got this code from someone else but I just extracted the required parts.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Beginner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-beginner.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-beginner-simplest-translucent-semi-transparent-form__1-60935/archive/master.zip)





### Source Code

```
Option Explicit
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Sub Form_Activate()
Dim NormalWindowStyle As Long
Dim HWD As Long
NormalWindowStyle = GetWindowLong(HWD, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes Me.hwnd, 0, --->(Any Value from 1 to 255)<--- , LWA_ALPHA
End Sub
```

