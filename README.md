<div align="center">

## Voice Recognition \(REALLY WORKS\-GREAT\!\)


</div>

### Description

Will Read Out Whatever You Type In!
 
### More Info
 
Insert MSAGENT control and name it Agent1

Insert TEXT BOX and name it Text1

Insert COMMAND BOX and name it Command1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Avison](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-avison.md)
**Level**          |Beginner
**User Rating**    |3.5 (46 globes from 13 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-avison-voice-recognition-really-works-great__1-33285/archive/master.zip)

### API Declarations

```
Const merlin = "merlin.acs"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
```


### Source Code

```
Private Sub Command1_Click()
 On Error Resume Next
     Set mer = Agent1.Characters("merlin")
     mer.LanguageID = &H409
     mer.Show
     mer.Stop
     mer.Speak Text1.Text
End Sub
Private Sub Form_Load()
  Agent1.Characters.Load "merlin", merlin
  Set mer = Agent1.Characters("merlin")
End Sub
```

