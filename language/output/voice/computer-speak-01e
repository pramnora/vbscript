DIM text(3) '...Create an array to hold lines of text.
            ' (NOTE: The array index number starts counting up from: 0.)

'*** Initialise each line of text inside of the array...

text(0)="Line 0"
text(1)="Line 1"
text(2)="Line 2"
text(3)="Line 3"

'*** Create a Windows Script object in order to get the computer to speak...

Set o=WScript.CreateObject("SAPI.SPVoice")

'*** Use a For/Next loop to traverse all the way through the text array...going from beginning to end.
'    By using LBound()/UBound() functions...;  
'    then, you don't need to keep on changing the For/Next loop size...ie. where does the array boundaries both begin/end(?);  
'    Because this will, automatically, be taken care of for you, instead.

For eachLine = LBound(text) to UBound(text)
 o.speak text(eachLine) '...speak each line of array text
Next
