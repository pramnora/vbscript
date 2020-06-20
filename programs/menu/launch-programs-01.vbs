'*** Created: 09:30 02/10/2017
'    Updated: 00:34 09/09/2019

strUnderlineChars=String(50,"*")

strTitleText="Please, enter a number..."

strMenuText="MENU: " & vbCrLf & vbCrLf &_
                   "<1> Notepad  - Word processor/(plain text ASCII)" & vbCrLf &_
                   "<2> Wordpad  - Word processor/(text/colours/graphics/-etc.)" & vbCrLf &_
                   "<3> MS Paint - Graphics(drawing/painting)" & vbCrLf &_
                   "<4>      CMD - Windows CLI/Command Line Interface" & vbCrLf & vbCrLf &_
                   "<5> YouTube  - inside of 'default' Web Browser" & vbCrLf & vbCrLf &_
                   strUnderlineChars & vbCrLf &_ 
                   "USER INSTRUCTIONS: " & vbCrLf & vbCrLf &_
                   "Launch a program by selecting a number either:" & vbCrLf &_
                   "<1>/<2>/<3>/<4>/<5>" & vbCrLf &_
                   "...and, then, pressing the ENTER key." & vbCrLf &_
                   "-Thanks!" 

strUserEntryText="Enter a number"

intUserNo=InputBox(strMenuText,strTitleText,strUserEntryText)

If intUserNo = 1 Then CreateObject("WScript.Shell").Run "notepad.exe"
If intUserNo = 2 Then CreateObject("WScript.Shell").Run "write.exe"
If intUserNo = 3 Then CreateObject("WScript.Shell").Run "mspaint.exe"
If intUserNo = 4 Then CreateObject("WScript.Shell").Run "cmd.exe"
If intUserNo = 5 Then CreateObject("WScript.Shell").Run "https://www.youtube.com"
