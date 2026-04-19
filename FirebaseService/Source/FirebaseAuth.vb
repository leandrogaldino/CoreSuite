Imports System.Net.Http
Imports System.Text
Imports CoreSuite.Helpers
''' <summary>
''' Provides authentication and session management for Firebase using email and password.
''' </summary>
''' <remarks>
''' <para>
''' The <see cref="FirebaseAuth"/> class handles user authentication against Firebase Authentication
''' services using the REST API. It manages access tokens, refresh tokens, and session expiration.
''' </para>
''' 
''' <para>
''' After a successful login, the class stores the authentication token in the associated
''' <c>FirebaseClient</c> instance and keeps track of its expiration time.
''' </para>
''' 
''' <para>
''' It also provides automatic session renewal using refresh tokens and ensures that API requests
''' can be executed with a valid authentication token.
''' </para>
''' 
''' <para>
''' This class is designed to be used internally by a <c>FirebaseClient</c> and should not be
''' instantiated directly outside of that context.
''' </para>
''' </remarks>
Public Class FirebaseAuth
    Private ReadOnly _Client As FirebaseClient
    Private _RefreshToken As String = String.Empty
    Private _TokenExpiration As DateTime = DateTime.MinValue
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Gets a value indicating whether the user is currently authenticated.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> if a valid authentication token is present; otherwise, <c>False</c>.
    ''' </returns>
    Public ReadOnly Property IsLoggedIn As Boolean
        Get
            Return Not String.IsNullOrEmpty(_Client.Token)
        End Get
    End Property
    ''' <summary>
    ''' Gets the current refresh token associated with the authenticated session.
    ''' </summary>
    ''' <returns>
    ''' The refresh token string if available; otherwise, an empty string.
    ''' </returns>
    Public ReadOnly Property RefreshToken As String
        Get
            Return _RefreshToken
        End Get
    End Property
    ''' <summary>
    ''' Authenticates the user with Firebase using email and password.
    ''' </summary>
    ''' <param name="Email">The user's email address.</param>
    ''' <param name="Password">The user's password.</param>
    ''' <returns>
    ''' A task that returns the refresh token if authentication is successful.
    ''' </returns>
    ''' <exception cref="Exception">
    ''' Thrown when authentication fails, there is no internet connection,
    ''' or the Firebase server does not respond.
    ''' </exception>
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
                        Dim seconds As Integer = If(Integer.TryParse(ExpiresIn, seconds), seconds, 3600)
                        _TokenExpiration = DateTime.UtcNow.AddSeconds(seconds - 60)
                        Return _RefreshToken
                    End If
                Else
                    Dim [Error] = TextHelper.ExtractJsonValue(ResponseBody, "message")
                    Throw New Exception($"Login failed: {[Error]}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("No internet connection.")
        Catch ex As TaskCanceledException
            Throw New Exception("The Google server took too long to respond.")
        End Try
        Return String.Empty
    End Function
    ''' <summary>
    ''' Attempts to refresh the current authentication session using the stored refresh token.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> if the session was successfully refreshed; otherwise, <c>False</c>.
    ''' </returns>
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
                    Dim Seconds As Integer = If(Integer.TryParse(ExpiresIn, Seconds), Seconds, 3600)
                    _TokenExpiration = DateTime.UtcNow.AddSeconds(Seconds - 60)
                    Return True
                Else
                    Debug.WriteLine($"Error refreshing token: {ResponseBody}")
                    Return False
                End If
            End Using
        Catch
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Ensures that the current authentication token is valid, refreshing it if necessary.
    ''' </summary>
    ''' <returns>
    ''' A task representing the asynchronous operation.
    ''' </returns>
    ''' <exception cref="Exception">
    ''' Thrown if the session has expired and cannot be refreshed.
    ''' </exception>
    Public Async Function EnsureValidTokenAsync() As Task
        If DateTime.UtcNow >= _TokenExpiration Then
            Dim Ok = Await RefreshSessionAsync()
            If Not Ok Then
                Throw New Exception("Session expired. Please log in again.")
            End If
        End If
    End Function
    Public Sub Logout()
        _Client.Token = Nothing
        _RefreshToken = Nothing
        _TokenExpiration = DateTime.MinValue
    End Sub
End Class