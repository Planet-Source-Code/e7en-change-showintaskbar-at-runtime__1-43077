<div align="center">

## Change ShowInTaskbar at Runtime


</div>

### Description

This code will allow you to change the .ShowInTaskBar at runtime. Please Comment and Vote.
 
### More Info
 
SetShowInTaskbar False, Me.hwnd or SetShowInTaskbar True, Me.hwnd


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ï¿½e7eN](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/e7en.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/e7en-change-showintaskbar-at-runtime__1-43077/archive/master.zip)





### Source Code

```
'Window API
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Window Constants
Public Const GWL_EXSTYLE = (-20)
Public Sub SetShowInTaskbar(Visible As Boolean, hWnd As Long)
  Dim Ret As Long
  Dim lNewLong As Long
  Select Case Visible
    Case True
      lNewLong = -128
    Case False
      lNewLong = 128
  End Select
  Ret = SetWindowLong(hWnd, GWL_EXSTYLE, lNewLong)
End Sub
```

