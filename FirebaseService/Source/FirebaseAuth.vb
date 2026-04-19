Imports System.Net.Http
Imports System.Text
Imports CoreSuite.Helpers

Public Class FirebaseAuth

    Private ReadOnly _Client As FirebaseClient
    Private _RefreshToken As String = String.Empty
    Private _TokenExpiration As DateTime = DateTime.MinValue

    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub

    Public ReadOnly Property IsLoggedIn As Boolean
        Get
            Return Not String.IsNullOrEmpty(_Client.Token)
        End Get
    End Property

    Public ReadOnly Property RefreshToken As String
        Get
            Return _RefreshToken
        End Get
    End Property

    ' =========================
    ' LOGIN
    ' =========================
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
                    Dim ExpiresIn = TextHelper.ExtractJsonValue(ResponseBody, "expiresIn")

                    If Not String.IsNullOrEmpty(Token) Then
                        _Client.Token = Token

                        ' controla expiração (com margem de segurança)
                        Dim seconds As Integer = If(Integer.TryParse(ExpiresIn, seconds), seconds, 3600)
                        _TokenExpiration = DateTime.UtcNow.AddSeconds(seconds - 60)

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
        End Try

        Return String.Empty
    End Function

    ' =========================
    ' REFRESH AUTOMÁTICO
    ' =========================
    Public Async Function RefreshSessionAsync() As Task(Of Boolean)
        If String.IsNullOrEmpty(_RefreshToken) Then Return False

        Try
            Dim Url = $"https://securetoken.googleapis.com/v1/token?key={_Client.ApiKey}"
            Dim Payload = $"grant_type=refresh_token&refresh_token={_RefreshToken}"

            Using Content As New StringContent(Payload, Encoding.UTF8, "application/x-www-form-urlencoded")
                Dim Response = Await _Client.Http.PostAsync(Url, Content)
                Dim ResponseBody = Await Response.Content.ReadAsStringAsync()

                If Response.IsSuccessStatusCode Then
                    _Client.Token = TextHelper.ExtractJsonValue(ResponseBody, "id_token")
                    _RefreshToken = TextHelper.ExtractJsonValue(ResponseBody, "refresh_token")

                    Dim ExpiresIn = TextHelper.ExtractJsonValue(ResponseBody, "expires_in")
                    Dim seconds As Integer = If(Integer.TryParse(ExpiresIn, seconds), seconds, 3600)

                    _TokenExpiration = DateTime.UtcNow.AddSeconds(seconds - 60)

                    Return True
                Else
                    Debug.WriteLine($"Erro ao renovar token: {ResponseBody}")
                    Return False
                End If
            End Using

        Catch
            Return False
        End Try
    End Function

    ' =========================
    ' GARANTE TOKEN VÁLIDO
    ' =========================
    Public Async Function EnsureValidTokenAsync() As Task
        If DateTime.UtcNow >= _TokenExpiration Then
            Dim ok = Await RefreshSessionAsync()
            If Not ok Then
                Throw New Exception("Sessão expirada. Faça login novamente.")
            End If
        End If
    End Function

    ' =========================
    ' LOGOUT
    ' =========================
    Public Sub Logout()
        _Client.Token = Nothing
        _RefreshToken = Nothing
        _TokenExpiration = DateTime.MinValue
    End Sub

End Class