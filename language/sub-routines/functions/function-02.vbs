'*** Here the function has values passed into it; 
'    does the calculation;
'    and, finally, returns a value...;
'    this value is, then, printed out 'outside' of the function itself.

MsgBox(add(5,4))

Function add(n1,n2)
 add = n1+n2
End Function

