'*** Here the function returns a value...;
'    which is, then, printed out 'outside' of the function itself.

MsgBox(add(5,4))

Function add(n1,n2)
 add = n1+n2
End Function

