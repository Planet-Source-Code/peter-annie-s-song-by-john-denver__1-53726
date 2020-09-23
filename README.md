<div align="center">

## Annie's Song by John Denver


</div>

### Description

No interface, just the song using Beep API... XP Only, sorry. I didn't write this but thought it might be enjoyed after seeing a single beep get 5 globes :p.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter.md)
**Level**          |Beginner
**User Rating**    |5.0 (104 globes from 21 users)
**Compatibility**  |VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-annie-s-song-by-john-denver__1-53726/archive/master.zip)





### Source Code

```
'    "Annie's Song"
'    by John Denver
Option Explicit
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_Load()
Dim Note As Long
Dim Frequencies As String, Durations As String
Dim Frequency As Long, Duration As Long
Frequencies = "iiihfihfffhidadddfhihfffhihiiihfihffihfdadddfhihffhiki"
 Durations = "aabbbfjaabbbbnaabbbfjaabcapaabbbfjaabbbbnaabbbfjaabcap"
Const E4 = 329.6276
For Note = 1 To Len(Frequencies)
  Frequency = E4 * 2 ^ ((Asc(Mid$(Frequencies, Note, 1)) - 96) / 12)
  Duration = (Asc(Mid$(Durations, Note, 1)) - 96) * 200 - 10
  Beep Frequency, Duration
  Sleep 10
  DoEvents
Next
Unload Me
End Sub
```

