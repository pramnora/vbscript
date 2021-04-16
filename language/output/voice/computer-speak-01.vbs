'*** Created: 08:30 14/04/2018
'    Updated: 08:30 14/04/2018
'    Code borrowed from...
'    https://www.youtube.com/watch?v=KCxerod4PBA
Dim strMsg,objSAPI
strMsg = "I the computer can speak!"
Set objSAPI = WScript.CreateObject("SAPI.Spvoice")
objSAPI.speak strMsg
