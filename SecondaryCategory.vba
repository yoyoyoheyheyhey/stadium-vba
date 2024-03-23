Private pName As String

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(Value As String)
    pName = Value
End Property

