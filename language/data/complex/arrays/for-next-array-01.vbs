'*** LBound(), captures the first array item number;
'    UBound(), captures the last array item number.

'    The array loop counts through each array item;
'    and, displays each array item 
'    using a standard windows output message box.     

Dim araItems(3)

araItems(0)="zero"
araItems(1)="one"
araItems(2)="two"
araItems(3)="three"

For intEachItemNo = LBound(araItems) To UBound(araItems)
 MsgBox(araItems(intEachItemNo))
Next 
