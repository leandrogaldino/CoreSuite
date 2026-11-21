Imports System.Runtime.Serialization
Namespace Global.System.Security.Cryptography
    <Serializable>
    Public Class AuthenticationTagMismatchException
        Inherits CryptographicException
        Public Sub New()
            MyBase.New("The authentication tag did not match.")
        End Sub
        Public Sub New(Message As String)
            MyBase.New(Message)
        End Sub
        Public Sub New(Message As String, InnerException As Exception)
            MyBase.New(Message, InnerException)
        End Sub
        Protected Sub New(Info As SerializationInfo, Context As StreamingContext)
            MyBase.New(Info, Context)
        End Sub
    End Class
End Namespace
