'*** Created: 10:15 19/12/2021
'    Updated: 10:15 19/12/2021
'    EXPLANATION: To encode/decode each pair of alphabet letters are swapped around: A -> B/B -> A/-etc.
strUserText = InputBox("Enter text to encode/decode: (a-z)/(A-Z)...","PROGRAM: ASCII Encode/Decode") 
For intEachChar = 1 TO Len(strUserText)
    strEachChar = Ucase(Mid(strUserText, intEachChar, 1))
    If strEachChar >= "A" And strEachChar <= "Z" Then
        intASCIINo = ASC(strEachChar) - 64
        If (intASCIINo) Mod 2 = 0 Then
            strEncDec = strEncDec + CHR((intASCIINo + 64) - 1)
        Else
            strEncDec = strEncDec + CHR((intASCIINo + 64) + 1)
        End If
    End If
Next
MsgBox("Here is your text encoded/decoded: " + vbCrLf + vbCrLf + strEncDec)
