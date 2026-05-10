Function CalcAndGreet(name, a, b)
    Dim total
    total = Add(a, b)
    CalcAndGreet = Greet(name) & " Sum=" & total
End Function
