'*** This for loop counts up from 1 to 10...;
'    and, prints out each even value loop number
'    inside of it's own separate message box.

For x = 1 TO 10
  If x MOD 2 = 0 Then 
 	 MsgBox(x)
  End If
Next

'...output...

'2
'4
'6
'8
'10
