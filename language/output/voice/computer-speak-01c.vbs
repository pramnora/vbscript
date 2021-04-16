Set o = WScript.CreateObject("SAPI.SPVoice")
xAnyText = Inputbox("Enter some text for the computer to talk...")
o.speak xAnyText
