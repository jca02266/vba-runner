Public Function CallPublic(a, b)
    CallPublic = PublicAdd(a, b)
End Function

Public Function CallPrivateCrossModule(x)
    CallPrivateCrossModule = PrivateHelper(x)
End Function
