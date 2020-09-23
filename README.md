<div align="center">

## Get the name of a control at runtime


</div>

### Description

I've recently taken over a project from someone else, and I've been left with code that has few naming conventions and a lot of bugs. I often find myself stepping through code wanting to check the value of a field. Unfortunately, I don't know the field's name-it could be Name, UserName, NameUser, txtName, and so on. It's a real pain to stop the program, click on the control in question, press [F4], get the control name, start the program again, and return to the point in the code where I was before. Here's a handy trick to get the control name right away.

by Jeff Brown; Jeff.Brown@piog.com; Pioneering Management Corporation
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |3.5 (28 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-get-the-name-of-a-control-at-runtime__1-639/archive/master.zip)





### Source Code

```
At runtime, set the focus to the control in question. Then, click the Break button on the VB toolbar, type
   ?Screen.ActiveControl.Name
in the debug window, and press [Enter]. Voila! VB displays the control name in the debug window-and you didn't have to stop the program.
```

