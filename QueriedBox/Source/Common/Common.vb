Public Class Common
    Friend Shared Sub ThrowNoTransparentColorException()
        Throw New ArgumentException("The control does not support transparent text colors.")
    End Sub
End Class
