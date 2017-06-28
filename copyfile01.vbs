'*** Code borrowed from...
'    http://www.vbsedit.com/html/3481a0bc-3fed-46cd-98a0-f3a53adfa0e1.asp

Dim FSOSet 
FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile "New Text Document.txt", "New Text Document 100.txt"
