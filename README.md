<div align="center">

## Change background/foreground color of Progressbar


</div>

### Description

Change background/foreground color of Progressbar.

using SENDMESSAGE/win32API
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Juha sÃƒÂ¯Ã‚Â¿Ã‚Â½derqvist](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/juha-s-derqvist.md)
**Level**          |Advanced
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/juha-s-derqvist-change-background-foreground-color-of-progressbar__1-33694/archive/master.zip)

### API Declarations

```
Public Declare Function SendMessage Lib _
 "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long
Public Const CCM_FIRST = &H2000
Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Public Const WM_USER = &H400
Public Const PBM_SETBARCOLOR = (WM_USER + 9)
```


### Source Code

```
' -------------- module code --------------
Public Declare Function SendMessage Lib _
 "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long
Public Const CCM_FIRST = &H2000
Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Public Const WM_USER = &H400
Public Const PBM_SETBARCOLOR = (WM_USER + 9)
' to change progressbarcolor
Public Sub colortoprogress(prog As Long, bgr As Integer, bgg As Integer, bgb As Integer, fgr As Integer, fgg As Integer, fgb As Integer)
SendMessage prog, PBM_SETBKCOLOR, 0, ByVal RGB(bgr, bgg, bgb)
SendMessage prog, PBM_SETBARCOLOR, 0, ByVal RGB(fgr, fgg, fgb)
End Sub
' -------------- form code --------------
Private Sub Form_Load()
Me.ProgressBar1.Scrolling = ccScrollingSmooth
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Max = 100
colortoprogress Me.ProgressBar1.hwnd, 255, 255, 255, 0, 0, 0
Timer1.Interval = 10
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
Me.ProgressBar1.visible = True
If Me.ProgressBar1.Value = 100 Then Me.ProgressBar1.Value = 1
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
End Sub
```

