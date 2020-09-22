Attribute VB_Name = "ModLeet"
Function Translate(TheText As String, FirstObj As String, SecondObj As String, Optional Reverse As Boolean = False) As String
    On Error Resume Next
        If Reverse = True Then
            Translate = Replace(TheText, SecondObj, FirstObj)
        Else
            Translate = Replace(TheText, FirstObj, SecondObj)
        End If
End Function

Public Function RunTranslate(RunOn As String, Optional Reverse As Boolean = False) As String
    On Error Resume Next
    '6 char
    RunOn = Translate(RunOn, "hacker", "haxx0r", Reverse)
    RunOn = Translate(RunOn, "please", "plz", Reverse)
    '4 char
    RunOn = Translate(RunOn, "hack", "hax", Reverse)
    '3 char
    RunOn = Translate(RunOn, "lol", "lawl", Reverse)
    RunOn = Translate(RunOn, "own", "pwn", Reverse)
    RunOn = Translate(RunOn, "the", "t3h", Reverse)
    '2 char
    RunOn = Translate(RunOn, "ex", "x", Reverse)
    '1 char
    RunOn = Translate(RunOn, "e", "3", Reverse)
    RunOn = Translate(RunOn, "A", "@", Reverse)
    RunOn = Translate(RunOn, "a", "4", Reverse)
    RunOn = Translate(RunOn, "i", "!", Reverse)
    RunOn = Translate(RunOn, "l", "1", Reverse)
    RunOn = Translate(RunOn, "S", "5", Reverse)
    RunOn = Translate(RunOn, "t", "7", Reverse)
    RunOn = Translate(RunOn, "T", "7", Reverse)
    RunOn = Translate(RunOn, "g", "9", Reverse)
    RunOn = Translate(RunOn, "z", "2", Reverse)
    RunOn = Translate(RunOn, "s", "z", Reverse)
    RunOn = Translate(RunOn, "o", "0", Reverse)
    RunOn = Translate(RunOn, "O", "0", Reverse)
    RunOn = Translate(RunOn, "N", "|\|", Reverse)
    RunOn = Translate(RunOn, "B", "8", Reverse)
    RunOn = Translate(RunOn, "M", "|\/|", Reverse)
    RunOn = Translate(RunOn, "w", "vv", Reverse)
    RunOn = Translate(RunOn, "W", "\/\/", Reverse)
    RunOn = Translate(RunOn, "H", "|-|", Reverse)
    RunOn = Translate(RunOn, "c", "(", Reverse)
    RunOn = Translate(RunOn, "C", "(", Reverse)
    RunOn = Translate(RunOn, "X", "}{", Reverse)
    RunOn = Translate(RunOn, "U", "(_)", Reverse)
    RunTranslate = RunOn
End Function

