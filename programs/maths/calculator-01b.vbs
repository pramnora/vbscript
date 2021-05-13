Do
 strSum = InputBox("Enter a sum (eg. 1+1)...")
 If strSum <> "" Then MsgBox(eval(strSum))
 strUserResponse = InputBox("Continue, Y/N?")
Loop Until UCase(strUserResponse) <> "Y"
