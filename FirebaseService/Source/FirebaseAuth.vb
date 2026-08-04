Imports System.Net.Http
Imports System.Text
Imports System.Text.Json
Imports System.Threading
''' <summary>
''' Authenticates Firebase users with email and password and manages their in-memory session.
''' </summary>
Public NotInheritable Class FirebaseAuth
    Private Shared ReadOnly RefreshThreshold As TimeSpan = TimeSpan.FromMinutes(1)
    Private ReadOnly _Client As FirebaseClient
    Private ReadOnly _RefreshLock As New SemaphoreSlim(1, 1)
    Private ReadOnly _SessionLock As New Object()
    Private _RefreshToken As String = String.Empty
    Private _TokenExpirationUtc As DateTime = DateTime.MinValue
    Private _UserId As String
    Private _Email As String
    Private _SessionVersion As Long
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Gets a value indicating whether a non-expired Firebase ID token is currently stored.
    ''' </summary>
    Public ReadOnly Property IsLoggedIn As Boolean
        Get
            SyncLock _SessionLock
                Return Not String.IsNullOrWhiteSpace(_Client.Token) AndAlso _TokenExpirationUtc > DateTime.UtcNow
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the current session has a refresh token.
    ''' </summary>
    Public ReadOnly Property CanRefreshSession As Boolean
        Get
            SyncLock _SessionLock
                Return Not String.IsNullOrWhiteSpace(_RefreshToken)
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets the Firebase user identifier returned by the latest successful sign-in or refresh operation.
    ''' </summary>
    Public ReadOnly Property UserId As String
        Get
            SyncLock _SessionLock
                Return _UserId
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets the email address returned by the latest successful sign-in operation.
    ''' </summary>
    Public ReadOnly Property Email As String
        Get
            SyncLock _SessionLock
                Return _Email
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets the UTC expiration time of the current Firebase ID token.
    ''' </summary>
    Public ReadOnly Property TokenExpirationUtc As DateTime
        Get
            SyncLock _SessionLock
                Return _TokenExpirationUtc
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Authenticates an existing Firebase user with an email address and password.
    ''' </summary>
    ''' <param name="Email">The user's email address.</param>
    ''' <param name="Password">The user's password.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <exception cref="ArgumentException">The email address or password is empty.</exception>
    ''' <exception cref="FirebaseException">Firebase rejects the credentials, returns an invalid response or cannot be reached.</exception>
    Public Async Function LoginAsync(Email As String, Password As String, Optional CancellationToken As CancellationToken = Nothing) As Task
        If String.IsNullOrWhiteSpace(Email) Then Throw New ArgumentException("The email address cannot be empty.", NameOf(Email))
        If String.IsNullOrEmpty(Password) Then Throw New ArgumentException("The password cannot be empty.", NameOf(Password))
        Dim Url As String = $"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={Uri.EscapeDataString(_Client.Options.ApiKey)}"
        Dim Payload As New Dictionary(Of String, Object) From {{"email", Email}, {"password", Password}, {"returnSecureToken", True}}
        Dim JsonBody As String = JsonSerializer.Serialize(Payload)
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Post, Url)
                    Request.Content = New StringContent(JsonBody, Encoding.UTF8, "application/json")
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Authentication, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Authentication, OperationSource.Token).ConfigureAwait(False)
                        Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                        Dim LoginResponse As FirebaseSignInResponse
                        Try
                            LoginResponse = JsonSerializer.Deserialize(Of FirebaseSignInResponse)(ResponseBody)
                        Catch ex As JsonException
                            Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Authentication, "Firebase Authentication returned malformed JSON.", ex)
                        End Try
                        ValidateSessionResponse(LoginResponse?.IdToken, LoginResponse?.RefreshToken, LoginResponse?.ExpiresIn)
                        ApplySession(LoginResponse.IdToken, LoginResponse.RefreshToken, LoginResponse.ExpiresIn, LoginResponse.LocalId, LoginResponse.Email)
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Authentication, ex)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Immediately exchanges the stored refresh token for a new Firebase ID token.
    ''' </summary>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns><see langword="True"/> when the session was refreshed; <see langword="False"/> when no refresh token is stored.</returns>
    ''' <exception cref="FirebaseException">Firebase rejects the refresh token, returns an invalid response or cannot be reached.</exception>
    Public Function RefreshSessionAsync(Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        Return RefreshSessionCoreAsync(True, CancellationToken)
    End Function
    ''' <summary>
    ''' Ensures that a valid Firebase ID token is available, refreshing it when it is near expiration.
    ''' </summary>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <exception cref="FirebaseException">No refreshable session exists or the refresh operation fails.</exception>
    Public Async Function EnsureValidTokenAsync(Optional CancellationToken As CancellationToken = Nothing) As Task
        If Not TokenNeedsRefresh() Then Return
        If Not Await RefreshSessionCoreAsync(False, CancellationToken).ConfigureAwait(False) Then Throw FirebaseException.CreateAuthenticationRequired()
    End Function
    ''' <summary>
    ''' Clears the ID token, refresh token and public session information held in memory.
    ''' </summary>
    Public Sub Logout()
        SyncLock _SessionLock
            _SessionVersion += 1
            _Client.Token = Nothing
            _RefreshToken = String.Empty
            _TokenExpirationUtc = DateTime.MinValue
            _UserId = Nothing
            _Email = Nothing
        End SyncLock
    End Sub
    Private Async Function RefreshSessionCoreAsync(ForceRefresh As Boolean, CancellationToken As CancellationToken) As Task(Of Boolean)
        Await _RefreshLock.WaitAsync(CancellationToken).ConfigureAwait(False)
        Try
            If Not ForceRefresh AndAlso Not TokenNeedsRefresh() Then Return True
            Dim CurrentRefreshToken As String
            Dim CurrentSessionVersion As Long
            SyncLock _SessionLock
                CurrentRefreshToken = _RefreshToken
                CurrentSessionVersion = _SessionVersion
            End SyncLock
            If String.IsNullOrWhiteSpace(CurrentRefreshToken) Then Return False
            Dim Url As String = $"https://securetoken.googleapis.com/v1/token?key={Uri.EscapeDataString(_Client.Options.ApiKey)}"
            Dim Values As New Dictionary(Of String, String) From {{"grant_type", "refresh_token"}, {"refresh_token", CurrentRefreshToken}}
            Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
                Try
                    Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Post, Url)
                        Request.Content = New FormUrlEncodedContent(Values)
                        Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Authentication, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                            Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Authentication, OperationSource.Token).ConfigureAwait(False)
                            Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                            Dim RefreshResponse As FirebaseRefreshResponse
                            Try
                                RefreshResponse = JsonSerializer.Deserialize(Of FirebaseRefreshResponse)(ResponseBody)
                            Catch ex As JsonException
                                Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Authentication, "Firebase Authentication returned malformed refresh JSON.", ex)
                            End Try
                            ValidateSessionResponse(RefreshResponse?.IdToken, RefreshResponse?.RefreshToken, RefreshResponse?.ExpiresIn)
                            Return TryApplyRefreshedSession(RefreshResponse.IdToken, RefreshResponse.RefreshToken, RefreshResponse.ExpiresIn, RefreshResponse.UserId, CurrentRefreshToken, CurrentSessionVersion)
                        End Using
                    End Using
                Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                    Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Authentication, ex)
                End Try
            End Using
        Finally
            _RefreshLock.Release()
        End Try
    End Function
    Private Function TokenNeedsRefresh() As Boolean
        SyncLock _SessionLock
            Return String.IsNullOrWhiteSpace(_Client.Token) OrElse _TokenExpirationUtc <= DateTime.UtcNow.Add(RefreshThreshold)
        End SyncLock
    End Function
    Private Sub ApplySession(IdToken As String, RefreshToken As String, ExpiresIn As String, UserId As String, Email As String)
        Dim Seconds As Integer = ParseExpirationSeconds(ExpiresIn)
        SyncLock _SessionLock
            _SessionVersion += 1
            _Client.Token = IdToken
            _RefreshToken = RefreshToken
            _TokenExpirationUtc = DateTime.UtcNow.AddSeconds(Seconds)
            If Not String.IsNullOrWhiteSpace(UserId) Then _UserId = UserId
            If Email IsNot Nothing Then _Email = Email
        End SyncLock
    End Sub
    Private Function TryApplyRefreshedSession(IdToken As String, RefreshToken As String, ExpiresIn As String, UserId As String, ExpectedRefreshToken As String, ExpectedSessionVersion As Long) As Boolean
        Dim Seconds As Integer = ParseExpirationSeconds(ExpiresIn)
        SyncLock _SessionLock
            If _SessionVersion <> ExpectedSessionVersion OrElse Not String.Equals(_RefreshToken, ExpectedRefreshToken, StringComparison.Ordinal) Then
                Return Not String.IsNullOrWhiteSpace(_Client.Token) AndAlso _TokenExpirationUtc > DateTime.UtcNow
            End If
            _SessionVersion += 1
            _Client.Token = IdToken
            _RefreshToken = RefreshToken
            _TokenExpirationUtc = DateTime.UtcNow.AddSeconds(Seconds)
            If Not String.IsNullOrWhiteSpace(UserId) Then _UserId = UserId
            Return True
        End SyncLock
    End Function
    Private Shared Sub ValidateSessionResponse(IdToken As String, RefreshToken As String, ExpiresIn As String)
        If String.IsNullOrWhiteSpace(IdToken) Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Authentication, "Firebase Authentication did not return an ID token.")
        If String.IsNullOrWhiteSpace(RefreshToken) Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Authentication, "Firebase Authentication did not return a refresh token.")
        ParseExpirationSeconds(ExpiresIn)
    End Sub
    Private Shared Function ParseExpirationSeconds(Value As String) As Integer
        Dim Seconds As Integer
        If Not Integer.TryParse(Value, Globalization.NumberStyles.None, Globalization.CultureInfo.InvariantCulture, Seconds) OrElse Seconds <= 0 Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Authentication, "Firebase Authentication returned an invalid token expiration value.")
        Return Seconds
    End Function
End Class
