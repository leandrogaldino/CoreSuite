Public Class CMessageBoxExceptionEmail
    Public Property FromName As String
    Public Property FromEmail As String
    Public Property ToName As String
    Public Property ToEmail As String
    Public Property Password As String
    Public Property Port As Integer
    Public Property Host As String
    Public Property SecureSocket As CMessageBoxSecureSocket = CMessageBoxSecureSocket.Auto
End Class
