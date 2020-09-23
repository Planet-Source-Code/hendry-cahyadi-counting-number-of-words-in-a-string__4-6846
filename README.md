<div align="center">

## Counting number of words in a string


</div>

### Description

This is a simple and easy way for counting the number of words in a string.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hendry Cahyadi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hendry-cahyadi.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hendry-cahyadi-counting-number-of-words-in-a-string__4-6846/archive/master.zip)





### Source Code

```
dim TheString, ArrayTemp, NumberOfWords, Word
TheString = "Hello, How are you today?" 'just a test string
ArrayTemp = split(TheString, " ")
NumberOfWords = UBound(ArrayTemp) + 1
Response.Write "<P>The string is: " & TheString
Response.Write "<P>Number of words in that string: " & NumberOfWords
Response.Write "<P>Here are the words which compose that string: "
for each Word in ArrayTemp
 Response.Write "<BR>" & word
next
```

