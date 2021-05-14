strTitleText = "PROGRAM: Calculator"

strBodyText = "Please, enter a sum:" &_
            + vbCrLf + vbCrLf + "3 + 3 (+ add/answer: 6)" &_
            + vbCrLf + vbCrLf + "3 - 3 (- subtract/answer: 0)" &_
            + vbCrLf + vbCrLf + "3 * 3 (* multiply/answer: 9)" &_
            + vbCrLf + vbCrLf + "3 / 3 (/ divide/answer: 1)" &_
            + vbCrLf + vbCrLf + "3 ^ 3 (^ exponentiation/answer: 27/3 x 3 x 3)" &_
            + vbCrLf + vbCrLf + "3 mod 5 (mod remainder/answer: 2)" &_
            + vbCrLf + vbCrLf + "------------------------------------------" &_
            + vbCrLf + vbCrLf + "NOTE: It is also possible to mix multiple different operators/numbers together: " &_
            + vbCrLf + vbCrLf + "3+5*7 (plus, multiply/answer: 38/5 x 7 = 35 + 3)" &_
            + vbCrLf + vbCrLf + "(3+5)*7 (plus, multiply/answer: 56/3 + 5 = 8 x 7)"

strInputBoxText = "1 + 1"

strSum = InputBox(strBodyText,strTitleText,strInputBoxText)
MsgBox(eval(strSum))
