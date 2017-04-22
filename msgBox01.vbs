'*** Created: 03:07 22/04/2017
'    Updated: 03:07 22/04/2017

'*** VBScript Message box template...
'    MsgBox "Message",Type,"Title"
'    Type: Code number: 0 = OK

x=MsgBox("Are you happy to go ahead?",vbYesNo+vbQuestion,"VBScript Message box to ask a question...")

If x=6 Then
 MsgBox("You Choose: Yes/Response code: (6)")
ElseIf x=7 Then
 MsgBox("You Choose: No/Response code: (7)")
End If
