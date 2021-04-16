Dim xAnyText
xAnyText = "Line 1" &_
           "Line 2" &_
           "Line 3"
Set o = WScript.CreateObject("SAPI.SPVoice")
o.speak xAnyText
