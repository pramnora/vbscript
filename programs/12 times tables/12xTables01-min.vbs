t = 0 
o = ""
n = InputBox("Enter a number times tables: ")
For t = 1 To 12
 o = o & t & " X " & n & " = " & t * n & vbCrLf 
Next
MsgBox(o)
