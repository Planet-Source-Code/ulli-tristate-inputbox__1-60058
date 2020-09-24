<div align="center">

## TriState InputBox


</div>

### Description

With the InputBox you cannot distinguish between the cases

a:- Cancel clicked

b:- nothing entered and OK clicked

because in both cases the returned string is a vbNullString

There is a simple trick however as is shown here
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Intermediate
**User Rating**    |4.9 (69 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ulli-tristate-inputbox__1-60058/archive/master.zip)





### Source Code

```
'With the InputBox you cannot distinguish between the cases
'
'  a:- Cancel clicked
'  b:- nothing entered and OK clicked
'
'because in both cases the returned string is a vbNullString
'
'There is a simple trick however as is shown in this little code snippet:
 Dim UserInput As String
  UserInput = InputBox("Please type in nothing or some text and click OK or Cancel", "Distinguish")
  Select Case True
   Case StrPtr(UserInput) = 0
    MsgBox "You clicked Cancel"
   Case Len(UserInput)
    MsgBox "You typed """ & UserInput & """ and clicked OK"
   Case Else
    MsgBox "You typed nothing and clicked OK"
  End Select
End Sub
```

