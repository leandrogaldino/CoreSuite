Imports Newtonsoft.Json
Friend NotInheritable Class FirebaseRefreshResponse
    <JsonProperty("id_token")>
    Public Property IdToken As String
    <JsonProperty("refresh_token")>
    Public Property RefreshToken As String
    <JsonProperty("expires_in")>
    Public Property ExpiresIn As String
    <JsonProperty("user_id")>
    Public Property UserId As String
End Class
