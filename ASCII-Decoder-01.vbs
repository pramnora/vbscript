'*** PROGRAM: ASCII Decoder...

'    CREATED: 08:48 23/04/2016
'    UPDATED: 08:48 23/04/2016 

'      USAGE: User is asked to enter their plain text message;
'             next, the program calculates...; then, outputs their message
'             as  being a series of ASCII encoded number pairs: (A-Z)/65-90. 

strEncodedText=InputBox("Enter encoded text...","PROGRAM: ASCII Decoder")

For intEachPairOfDigits = 1 To Len(strEncodedText) Step 2
        intASCIICode=Eval(Mid(strEncodedText,intEachPairOfDigits,2))
        If intASCIICode >=65 And intASCIICode <=90 Then
		strPlainText = strPlainText + Chr(intASCIICode)
	End if
Next

msgBox(strPlainText)
