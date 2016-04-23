'*** PROGRAM: ASCII Decoder...

'    CREATED: 08:48 23/04/2016
'    UPDATED: 08:48 23/04/2016 

'      USAGE: First, the user is asked to enter their ASCII encoded ciphertext message.
'             Next, the program decodes each pair of ASCII encoded digits: (65),
'             by changing it to become a single alphabet character, instead: (A); 
'             after decoding all ASCII encoded characters...the program outputs the user's ciphertext
'             as being a full plain text message consisting of letters: A-Z/(ASCII: 65-90). 

strEncodedText=InputBox("Enter encoded text...","PROGRAM: ASCII Decoder")

For intEachPairOfDigits = 1 To Len(strEncodedText) Step 2
        intASCIICode=Eval(Mid(strEncodedText,intEachPairOfDigits,2))
        If intASCIICode >=65 And intASCIICode <=90 Then
		strPlainText = strPlainText + Chr(intASCIICode)
	End if
Next

msgBox(strPlainText)
