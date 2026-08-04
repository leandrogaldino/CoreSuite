Imports System.Text.Json.Serialization
Friend NotInheritable Class FirebaseRefreshResponse
    <JsonPropertyName("id_token")>
    Public Property IdToken As String
    <JsonPropertyName("refresh_token")>
    Public Property RefreshToken As String
    <JsonPropertyName("expires_in")>
    Public Property ExpiresIn As String
    <JsonPropertyName("user_id")>
    Public Property UserId As String
End Class
