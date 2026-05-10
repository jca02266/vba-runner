Public Function PublicAdd(a, b)
    PublicAdd = a + b
End Function

Private Function PrivateHelper(x)
    PrivateHelper = x * 2
End Function

Public Function UseOwnPrivate(x)
    UseOwnPrivate = PrivateHelper(x)
End Function
