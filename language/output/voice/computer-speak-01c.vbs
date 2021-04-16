xAnyText = Inputbox("Enter some text for the computer to talk...")
Set o = WScript.CreateObject("SAPI.SPVoice")
o.speak xAnyText
