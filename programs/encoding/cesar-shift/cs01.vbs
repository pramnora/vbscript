'---------------------------------------------
'***  PROGRAM: Text encoding/Ceaser Shift
'    LANGUAGE: VBScript/Visual BASIC Script

'    LOCATION: WMCollege, Mornington Crescent branch, London, UK.
'          OS: Windows 10 Enterprise
'    COMPUTER: PC

'    CREATED: Thu 21st March 2024 11:37 AM GMT
'    UPDATED: Thu 21st March 2024 11:37 AM GMT
'---------------------------------------------

'------------------------------------------------------------------------

'*** Explanation: 

'    This program will encode any text message typed in...
'    by using a Caesar Shift cipher key of: +1.

'    -(The program will, automatically, change any lower case letters...
'      into being upper case letters, instead)-

'    This means:
'    A is translated to become B
'    B becomes C
'    C becomes D
'    -etc.
'    ...and, when the last letter in the alphabet: 'Z' is reached
'    it is translated as 'A'.

'-------------------------------------------------------------------------

'*** Initialise variables...

cesarShift=1

'-------------------
'*** Main program...
'-------------------

'*** Get user text.../
'    changing all lower case letters into being upper case, instead

userText=UCase(inputBox("Enter text"))

'*** Use loop to encode text...

for eachChar=1 to Len(userText)

 '*** get each letter in user text...
 aLetter=Mid(userText,eachChar,1)

 '*** Check if it's an alphabet letter, first, before doing any encoding...
 If aLetter >="A" And aLetter <="Z" Then

  '*** encode each letter in user text...

  If aLetter="Z" Then
   encodedText=encodedText+"A"  
  Else
   encodedText=encodedText + Chr(Asc(aLetter)+cesarShift)
  End If

 End If

next

'*** Show encoded text...

msgBox(encodedText)

'------------------------
'Example Run(1):-

'Enter text: abc
'    Output: BCD
'------------------------
'Example Run(2):-

'Enter text: This is a secret message.
'    Output: UIJTJTBTFDSFUNFTTBHF
'------------------------
