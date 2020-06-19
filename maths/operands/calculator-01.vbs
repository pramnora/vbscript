'-----------------------------
'*** PROGRAM: Calculator...

'    Created: 23:28 19/06/2020
'    Updated: 23:28 19/06/2020
'-----------------------------

'----------------------------
'*** Variable declarations...
'----------------------------

n1 = 5 : n2 = 4 
out = ""

'-------------------
'*** Main program...
'-------------------

out = "n1,n2=" + CStr(n1) + "," + CStr(n2) + vbCRLF
out = out + "a,s,m,d" + vbCRLF
out = out + CStr(add(n1,n2) & ",")
out = out + CStr(subtract(n1,n2) & ",")
out = out + CStr(multiply(n1,n2) & ",")
out = out + CStr(divide(n1,n2))

msgBox(out)
 
'---------------
'*** Functions...
'----------------

Function add(a,b)
	add = a+b
End Function

Function subtract(a,b)
 	subtract = a-b
End Function

Function multiply(a,b)
 	multiply = a*b
End Function

Function divide(a,b)
	divide = a/b
End Function

