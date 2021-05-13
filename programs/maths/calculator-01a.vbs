'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    PROGRAM: Calculate Sum
'   LANGUAGE: VBScript/Visual BASIC Script
'-----------------------------------------------------------
'   COMPUTER: PC/Windows 10 Professional
'     EDITOR: Notepad
'-----------------------------------------------------------
'  COPYRIGHT: 2021-, Mr. Paul Ramnora./All rights reserved.
'-----------------------------------------------------------
'    CREATED: 13:33 13/05/2021
'    UPDATED: 13:34 13/05/2021
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'*** WHAT THE PROGRAM DOES...

' A> This VBScript program asks the user to type in a number sum...
'    ...the sum may be as simple as: 1+1
'    ...or, more complex: 13+265/3*20/-etc.

' B> Next, the program will evaluate the sum the user has typed in;
'    and, print out the sum result.

' C> Finally, the program asks the user to type in a response either: 'Y/N';
'    if the user response typed is not a 'Y'; then, the program ends/
'    otherwise, the program continues by starting again from scratch.  

'-------------------
'*** Main Program...
'-------------------

Do '--Start loop...

 '*** Get user to enter a sum...

  strSum = InputBox("Enter a sum (eg. 1+1)...") '-- show an input box asking user to type in a sum

 '*** do calculation...

  If strSum <> "" Then   '-- start if block/check if sum value isn't empty

    MsgBox(eval(strSum)) '-- calculate/then, show sum result inside of output message box

 End If              '-- end if block

 '*** do re-Run...

  strUserResponse = InputBox("Continue, Y/N?") '-- show an input box asking user if they wish to continue or not

Loop Until UCase(strUserResponse) <> "Y" '-- end loop
