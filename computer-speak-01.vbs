'*** Created: 08:30 14/04/2018
'    Updated: 08:30 14/04/2018
'    Code borrowed from...
'    https://www.youtube.com/watch?v=KCxerod4PBA
Dim msg
msg = "I the computer can speak!"
Set objMyComputer = WScript.CreateObject("SAPI.Spvoice")
objMyComputer.speak msg
