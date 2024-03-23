Private pName As String
Private pChildren As Collection

Private Sub Class_Initialize()
    Set pChildren = New Collection
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(Value As String)
    pName = Value
End Property

Public Property Get Children() As Collection
    Set Children = pChildren
End Property

Public Sub AddChild(child As SecondaryCategory)
    pChildren.Add child
End Sub

