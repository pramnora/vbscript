'*** Created: 03:07 22/04/2017
'    Updated: 03:07 22/04/2017

'*** Code borrowed from YouTube video...
'    VBScript Basics, Part 1 | Message Box - Numbers (MsgBox)/(Tutorials by: SimplyCoded) 
'    https://www.youtube.com/watch?v=NwIOuZZqolE&list=PL72Es31dJnK6-ZFXXHFurNwJ6QtWPtxbz

'*** VBScript Message box template...
'    MsgBox "Message",Type,"Title"
'    Type: Code number: 0 = OK

x=MsgBox("Are you happy to go ahead?",vbYesNo+vbQuestion,"VBScript Message box to ask a question...")

If x=6 Then
 MsgBox("You Choose: Yes/Response code: (6)")
ElseIf x=7 Then
 MsgBox("You Choose: No/Response code: (7)")
End If
