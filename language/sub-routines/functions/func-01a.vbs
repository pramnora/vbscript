'*** This function adds 2 numbers;
'    and, then, outputs their sum result
'    into a standard windows message box...

MsgBox(add(1,2))  '...function call/
                  '   call function by its name,
                  '   in order to identify which function we are calling;
                  '   with a comma separated arguments list
                  '   being written inside a pair of parenthesis: ();
                  '   the arguments list should 'match' 
                  '   the named function's parameter list

FUNCTION add(x,y) '...function header/function start...
                  '   with comma separated parameter list;
                  '   the comma separated parameter list 
                  '   is used to receive the function call arguments

	add=x+y         '   function body/
                  '   use function name
                  '   together with an equals sign: =;
                  '   to return the sum result
                  '   back to whereever the function was called

END FUNCTION      '   ...function tail/function end

'*** In reality, this code is just 4 lines long; 
'    however, the comments make it appear to be far longer than in actually is;
'    the comment lines all begin with a single apostrophe: (');
'    if you were to remove all of the comment lines...;
'    then, this code would considerably shrink down in size.
'    Please, view the file: func-02.vbs;
'    to see this same program haveing been shrunk down.
