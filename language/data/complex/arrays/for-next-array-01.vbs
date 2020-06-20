'    The array loop explicitly counts up through each array item;
'    going from item: 0...all the way up to...item: 3;
'    which represents both the start/and, end array item numbers.
'    At the same time, it displays each array item 
'    inside of a standard windows output message box.     

Dim araItems(3)

araItems(0)="zero"
araItems(1)="one"
araItems(2)="two"
araItems(3)="three"

For intEachItemNo = 0 To 3
 MsgBox(araItems(intEachItemNo))
Next 
