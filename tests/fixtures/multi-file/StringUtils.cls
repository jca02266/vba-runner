Function Greet(name)
    Greet = "Hello, " & name & "!"
End Function

Function Repeat(s, n)
    Dim result
    result = ""
    Dim i
    For i = 1 To n
        result = result & s
    Next i
    Repeat = result
End Function
