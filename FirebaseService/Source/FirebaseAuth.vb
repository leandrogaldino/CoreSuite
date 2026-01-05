Imports System.Net.Http
Imports System.Text
Imports CoreSuite.Helpers
''' <summary>
''' Provides Firebase Authentication features such as login, session refresh, and logout.
''' </summary>
Public Class FirebaseAuth
    Private ReadOnly _Client As FirebaseClient
    Private _RefreshToken As String = String.Empty
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Indicates whether a user is currently authenticated.
    ''' </summary>
    Public ReadOnly Property IsLoggedIn As Boolean
        Get
            Return Not String.IsNullOrEmpty(_Client.Token)
        End Get
    End Property
    ''' <summary>
    ''' Gets the current refresh token.
    ''' </summary>
    Public ReadOnly Property RefreshToken As String
        Get
            Return _RefreshToken
        End Get
    End Property
    ''' <summary>
    ''' Authenticates a user using email and password.
    ''' </summary>
    ''' <param name="Email">User email address.</param>
    ''' <param name="Password">User password.</param>
    ''' <returns>The refresh token if authentication succeeds.</returns>
    Public Async Function LoginAsync(Email As String, Password As String) As Task(Of String)
        Try
            Dim Url = $"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={_Client.ApiKey}"
            Dim JsonBody = "{" &
                """email"": """ & Email & """," &
                """password"": """ & Password & """," &
                """returnSecureToken"": true" &
            "}"
            Using Content As New StringContent(JsonBody, Encoding.UTF8, "application/json")
                Dim Response = Await _Client.Http.PostAsync(Url, Content)
                Dim ResponseBody = Await Response.Content.ReadAsStringAsync()
                If Response.IsSuccessStatusCode Then
                    Dim Token = TextHelper.ExtractJsonValue(ResponseBody, "idToken")
                    _RefreshToken = TextHelper.ExtractJsonValue(ResponseBody, "refreshToken")
                    If Not String.IsNullOrEmpty(Token) Then
                        _Client.Token = Token
                        Return _RefreshToken
                    End If
                Else
                    Dim [Error] = TextHelper.ExtractJsonValue(ResponseBody, "message")
                    Throw New Exception($"Falha no login: {[Error]}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Sem conexão com a internet.")
        Catch ex As TaskCanceledException
            Throw New Exception("O servidor do Google demorou a responder.")
        Catch ex As Exception
            Throw
        End Try
        Return String.Empty
    End Function
    ''' <summary>
    ''' Renews the authentication session using a refresh token.
    ''' </summary>
    ''' <param name="RefreshToken">Refresh token previously issued by Firebase.</param>
    ''' <returns>True if the session was successfully refreshed.</returns>
    Public Async Function RefreshSessionAsync(RefreshToken As String) As Task(Of Boolean)
        Try
            Dim Url = $"https://securetoken.googleapis.com/v1/token?key={_Client.ApiKey}"
            Dim Payload = $"grant_type=refresh_token&refresh_token={RefreshToken}"
            Using Content As New StringContent(Payload, Encoding.UTF8, "application/x-www-form-urlencoded")
                Dim Response = Await _Client.Http.PostAsync(Url, Content)
                If Response.IsSuccessStatusCode Then
                    Dim ResponseBody = Await Response.Content.ReadAsStringAsync()
                    _Client.Token = TextHelper.ExtractJsonValue(ResponseBody, "id_token")
                    Return True
                Else
                    Dim [Error] = Await Response.Content.ReadAsStringAsync()
                    Debug.WriteLine($"Erro de Autenticação: {[Error]}")
                    Return False
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Erro de rede ao renovar sessão. Tente novamente.")
        Catch ex As TaskCanceledException
            Throw New Exception("Servidor demorou a responder.")
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Logs out the current user and clears the authentication token.
    ''' </summary>
    Public Sub Logout()
        _Client.Token = Nothing
    End Sub
End Class