Imports Newtonsoft.Json
Friend NotInheritable Class FirebaseSignInResponse
    <JsonProperty("idToken")>
    Public Property IdToken As String
    <JsonProperty("refreshToken")>
    Public Property RefreshToken As String
    <JsonProperty("expiresIn")>
    Public Property ExpiresIn As String
    <JsonProperty("localId")>
    Public Property LocalId As String
    <JsonProperty("email")>
    Public Property Email As String
End Class
