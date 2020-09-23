<div align="center">

## Filter numbers or characters from sting


</div>

### Description

Filter a string and return numbers or characters
 
### More Info
 
only numbers or only characters and optional the filtered numbers and characters in textbox


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Richard\_W](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/richard-w.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/richard-w-filter-numbers-or-characters-from-sting__1-44239/archive/master.zip)





### Source Code

```
Public Function NumberOrNoNumber(StrToCheck As String, Numbers As Boolean, Optional NumericTextTarget As TextBox, Optional TextualTextTarget As TextBox)
'example:
'    txtFilter = NumberOrNoNumber(txtStringIncludingNumbers, False, txtNumber, txtNoNumber)
'    txtFilter = NumberOrNoNumber(txtStringIncludingNumbers, True)
Dim Nstr As String 'targetstring for al numbers
Dim Tstr As String 'targetstring for everything exept numbers
Dim i As Integer
  For i = 1 To Len(StrToCheck)
    If IsNumeric(Mid(StrToCheck, i, 1)) Then Nstr = Nstr & Mid(StrToCheck, i, 1) Else Tstr = Tstr & Mid(StrToCheck, i, 1)
  Next
If Numbers Then NumberOrNoNumber = Nstr Else NumberOrNoNumber = Tstr
On Error Resume Next
NumericTextTarget = Nstr 'optional target for the numbers filtered out
TextualTextTarget = Tstr 'optional target for the text filtered out
End Function
```

