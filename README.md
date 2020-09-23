<div align="center">

## Exclusive Mode \(Updated\)


</div>

### Description

This function allows the application to enter and exit exclusive mode. In this mode any message boxes or prompts from Windows and other applications will not show up infront of the program. This is useful when you don't want anything to come up infront of your application window.
 
### More Info
 
Reguires a True to turn exclusive mode on and False to turn it off.

Essentially this code makes Windows think that your application is a screen saver. The only type of application that Windows will not interupt with message boxes, etc.

This code may cause some problems with screen savers that do not use the normal Windows interface such as After Dark. Users should be made aware of this.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Nawara](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-nawara.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-nawara-exclusive-mode-updated__1-762/archive/master.zip)

### API Declarations

```
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
```


### Source Code

```
Public Sub Exclusive_Mode(Use As Boolean)
'If True was passed makes app exclusive
'Else makes app not exclusive
Dim Scrap
Scrap = SystemParametersInfo(97, Use, "", 0)
End Sub
```

