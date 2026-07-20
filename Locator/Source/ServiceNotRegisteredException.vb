Public Class ServiceNotRegisteredException
    Inherits InvalidOperationException
    Public Sub New(Message As String)
        MyBase.New(Message)
    End Sub
End Class
