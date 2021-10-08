'---------------------------------------------------
'*** PROGRAM: 12 X Tables...
'   LANGUAGE: VBScript/Visual BASIC Script
'   COMPUTER: Home based PC/Windows 10 Professional
'     EDITOR: Notepad
'---------------------------------------------------
'     AUTHOR: Paul Ramnora
'      EMAIL: pramnora@yahoo.com
'  COPYRIGHT: Mr. Paul Ramnora./All rights reserved.
'---------------------------------------------------
'    CREATED: 22:25 08/10/2021
'    UPDATED: 22:25 08/10/2021 
'---------------------------------------------------

'----------------------------
'*** Variable declarations...
'----------------------------

intTablesNo = 0      '...loop counter variable
strOutString = ""    '...string, used to store each loop calculation output

'-------------------
'*** Main program...
'-------------------

'*** Get user to input a times tables number from the keyboard... 

intTablesNo = InputBox("Enter a number times tables: ")

'*** Do calculations.../and, store printout in outString...

For intTimesNo = 1 To 12 '...For/Next loop counts up from 1 to 12

 strOutString = strOutString &_
             intTimesNo & " X " & intTablesNo & " = " &_
             intTimesNo * intTablesNo &_
             vbCrLf 

Next

'*** Show times tables as output...

MsgBox(strOutString)
