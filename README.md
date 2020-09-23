<div align="center">

## Count number ocurrences of a sequence characters in string \* YOU MUST SEE \* THIS OWN\!\!\! SEE IT\!


</div>

### Description

Count number ocurrences of a sequence characters in string
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marco Antonio Alves](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marco-antonio-alves.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marco-antonio-alves-count-number-ocurrences-of-a-sequence-characters-in-string-you-must-se__1-28072/archive/master.zip)





### Source Code

```
Function NumChars(strSource, strSearchString) As Long
if len(strsource) then
 NumChars = UBound(Split(strSource, strSearchString))
end if
End Function
```

