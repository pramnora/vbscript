'---------------------------------------------
'***  PROGRAM: Text encoding/Ceaser Shift
'    LANGUAGE: VBScript/Visual BASIC Script

'    CREATED: Thu 21st March 2024 11:37 AM GMT
'    UPDATED: Thu 21st March 2024 11:37 AM GMT
'---------------------------------------------
     
'*** Initialise variables...

cesarShift=1

'-------------------
'*** Main program...
'-------------------

'*** Get user text...

userText=inputBox("Enter text")

'*** Use loop to encode text...

for eachChar=1 to Len(userText)

 '*** get each letter in user text...
 aLetter=Mid(userText,eachChar,1)

 '*** encode each letter in user text...
 encodedText=encodedText + Chr(Asc(aLetter)+cesarShift)

next

'*** Show encoded text...

msgBox(encodedText)
