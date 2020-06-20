'*** The line: 'Randomize Timer' is necesary for the numbers shown to be different in sequence;
'    without the numbers thrown will always be the same.          

Randomize Timer
intDiceNo = Int(Rnd*6)+1
MsgBox(intDiceNo)
