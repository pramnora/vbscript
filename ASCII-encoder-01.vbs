'*** PROGRAM: ASCII Encoder...
'    CREATED: 08:29 23/04/2016
'    UPDATED: 08:29 23/04/2016 

'      USAGE: First, the user is asked to enter their plain text message;
'             next, the program calculates...; then, outputs their message
'             as  being a series of ASCII encoded number pairs: (A-Z)/65-90. 

strPlainText=InputBox("Enter some plain text...","PROGRAM: ASCII Encoder")

For intEachChar = 1 To Len(strPlainText)
        strEachChar=Ucase(Mid(strPlainText,intEachChar,1))
        If strEachChar >="A" And strEachChar <="Z" Then
		strCipherText = strCipherText & Asc(strEachChar)
	End if
Next

msgBox(strCipherText)
