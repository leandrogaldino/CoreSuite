Imports System.Text.Json.Serialization
Friend NotInheritable Class FirebaseSignInResponse
    <JsonPropertyName("idToken")>
    Public Property IdToken As String
    <JsonPropertyName("refreshToken")>
    Public Property RefreshToken As String
    <JsonPropertyName("expiresIn")>
    Public Property ExpiresIn As String
    <JsonPropertyName("localId")>
    Public Property LocalId As String
    <JsonPropertyName("email")>
    Public Property Email As String
End Class
