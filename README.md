<div align="center">

## A Flashing Form


</div>

### Description

This code causes the form's TitleBar and the form's TaskBar Window to Flash.
 
### More Info
 
Create a Timer control and name it - Timer1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tomer Cohen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tomer-cohen.md)
**Level**          |Unknown
**User Rating**    |4.2 (79 globes from 19 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tomer-cohen-a-flashing-form__1-4842/archive/master.zip)

### API Declarations

```
Private Declare Function FlashWindow Lib "user32" Alias "FlashWindow" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
```


### Source Code

```
Private Sub Form_Load()
 Timer1.Interval = 300 'Change value depending on the speed of flahing.
End Sub
Private Sub Timer1_Timer()
 FlashWindow hwnd, 1
End Sub
```

