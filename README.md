<div align="center">

## A Must See Tips: Real Numeric Character in TextBox\. Avoid Paste Alpha Character


</div>

### Description

Many tips (and trick) tell us that if we want the textbox control ignore the character which is not numeric character, then we can just put the code that shown in KeyPress event procedure below. But, sometimes we forgot that although we have put the code in KeyPress event procedure, user still can input the alpha character to textbox control by doing copy and Paste to textbox. So, here is another tips to fix the problem. I hope this helpful.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Masino Sinaga](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/masino-sinaga.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/masino-sinaga-a-must-see-tips-real-numeric-character-in-textbox-avoid-paste-alpha-characte__1-47381/archive/master.zip)





### Source Code

```
Private Sub Text1_KeyPress(KeyAscii As Integer)
 If Not (KeyAscii >= Asc("0") & Chr(13) _
   And KeyAscii <= Asc("9") & Chr(13) _
   Or KeyAscii = vbKeyBack _
   Or KeyAscii = vbKeyDelete _
   Or KeyAscii = vbKeySpace) Then
    Beep
    KeyAscii = 0
  End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
 If Not (KeyAscii >= Asc("0") & Chr(13) _
   And KeyAscii <= Asc("9") & Chr(13) _
   Or KeyAscii = vbKeyBack _
   Or KeyAscii = vbKeyDelete _
   Or KeyAscii = vbKeySpace) Then
    Beep
    KeyAscii = 0
  End If
End Sub
'If user paste the character which is not
'numeric character, Text1 will ignore it.
Private Sub Text1_Change()
 If Not IsNumeric(Text1.Text) Then
   Text1.Text = ""
 End If
End Sub
'Try Paste some character which is not numeric
'to Text1 and Text2 control (copy alpha character
'from another file, paste it to those textboxes).
'See the difference between Text1 and Text2!!!
'So, don't forget to add the code in event
'procedure Change belongs to the textbox if
'you want your textbox control avoid the character
'which is not numeric. This is often we forgot!
```

