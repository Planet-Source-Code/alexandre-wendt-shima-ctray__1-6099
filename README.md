<div align="center">

## CTray


</div>

### Description

With CTray you can add/remove/change one icon to the System Tray. All you need to do is add CTray to your project and one picture box to your form, and you're done! A sample project is included.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-02-18 01:50:54
**By**             |[Alexandre Wendt Shima](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alexandre-wendt-shima.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD34502172000\.zip](https://github.com/Planet-Source-Code/alexandre-wendt-shima-ctray__1-6099/archive/master.zip)

### API Declarations

```
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNid As NOTIFYICONDATA) As Boolean
```





