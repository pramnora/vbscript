'*** The line: 'Randomize Timer', is necessary for the random numbers thrown to be 'different' in sequence;
'    without it, the dice numbers thrown will always follow 'the same' sequence.          

Randomize Timer
intDiceNo = Int(Rnd*6)+1
MsgBox(intDiceNo)
