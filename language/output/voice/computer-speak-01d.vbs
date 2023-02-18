'*** This method is using single string concatenation...; 
'    together with the VBScript line continuation character: '&_'

Dim xAnyText
xAnyText = "Line 1" &_
           "Line 2" &_
           "Line 3"
Set o = WScript.CreateObject("SAPI.SPVoice")
o.speak xAnyText
