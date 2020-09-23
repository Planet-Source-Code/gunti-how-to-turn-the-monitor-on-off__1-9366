<div align="center">

## How to turn the monitor on/off


</div>

### Description

With this code you can turn on or off the monitor ;)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[gunti](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gunti.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gunti-how-to-turn-the-monitor-on-off__1-9366/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Private Const WM_SYSCOMMAND = &H112
 Private Const SC_MONITORPOWER = &HF170
```


### Source Code

```
'Turn Monitor on: ->
 SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 0&
'Turn Monitor off: ->
 SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal -1&
```

