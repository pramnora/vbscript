'*** PROGRAM: ASCII Encoder...
'    CREATED: 08:29 23/04/2016
'    UPDATED: 08:29 23/04/2016 

'      USAGE: First, the user is asked to enter their 'plaintext' message.
'             Next, the program changes each plain english letter: (A),
'             to become an ASCII encoded digit pair of numbers, instead: (65).
'             Finally, the program outputs the 'plaintext' message...as being...'cipertext'/
'             which is a series of ASCII encoded number pairs: 65-90/(Alphabet: A-Z). 

strPlainText=InputBox("Enter some plain text...","PROGRAM: ASCII Encoder")

For intEachChar = 1 To Len(strPlainText)
        strEachChar=Ucase(Mid(strPlainText,intEachChar,1))
        If strEachChar >="A" And strEachChar <="Z" Then
		strCipherText = strCipherText & Asc(strEachChar)
	End if
Next

msgBox(strCipherText)

'*** EXAMPLE output...
'Enter some plain text...
'ABC
'656667
