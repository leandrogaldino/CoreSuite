Imports System.Net.Http
''' <summary>
''' Provides one configured entry point for Firebase Authentication, Cloud Firestore and Cloud Storage.
''' </summary>
''' <remarks>
''' All exposed modules share the same HTTP client and authenticated session. Dispose the service when it creates its own HTTP client.
''' </remarks>
Public NotInheritable Class FirebaseService
    Implements IDisposable
    Private ReadOnly _SyncRoot As New Object()
    Private _Client As FirebaseClient
    Private _Auth As FirebaseAuth
    Private _Firestore As FirebaseFirestore
    Private _Storage As FirebaseStorage
    Private _Disposed As Boolean
    ''' <summary>
    ''' Gets a value indicating whether the service has been initialized.
    ''' </summary>
    Public ReadOnly Property IsInitialized As Boolean
        Get
            SyncLock _SyncRoot
                Return Not _Disposed AndAlso _Client IsNot Nothing
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets a copy of the active Firebase configuration.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">The service has not been initialized.</exception>
    Public ReadOnly Property Options As FirebaseOptions
        Get
            SyncLock _SyncRoot
                Return GetRequiredClient().Options.Copy()
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets the Firebase Authentication module.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">The service has not been initialized.</exception>
    Public ReadOnly Property Auth As FirebaseAuth
        Get
            SyncLock _SyncRoot
                GetRequiredClient()
                Return _Auth
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets the Cloud Firestore module.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">The service has not been initialized.</exception>
    Public ReadOnly Property Firestore As FirebaseFirestore
        Get
            SyncLock _SyncRoot
                GetRequiredClient()
                Return _Firestore
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Gets the Cloud Storage module.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">The service has not been initialized.</exception>
    Public ReadOnly Property Storage As FirebaseStorage
        Get
            SyncLock _SyncRoot
                GetRequiredClient()
                Return _Storage
            End SyncLock
        End Get
    End Property
    ''' <summary>
    ''' Initializes an unconfigured service. Call <see cref="Initialize(FirebaseOptions)"/> before accessing a module.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a service with the required Firebase project values.
    ''' </summary>
    ''' <param name="ApiKey">The Firebase Web API key.</param>
    ''' <param name="ProjectId">The Firebase project identifier.</param>
    ''' <param name="StorageBucket">The Cloud Storage bucket name.</param>
    Public Sub New(ApiKey As String, ProjectId As String, StorageBucket As String)
        Initialize(New FirebaseOptions(ApiKey, ProjectId, StorageBucket))
    End Sub
    ''' <summary>
    ''' Initializes a service with a configuration object and an internally managed HTTP client.
    ''' </summary>
    ''' <param name="Options">The Firebase configuration.</param>
    Public Sub New(Options As FirebaseOptions)
        Initialize(Options)
    End Sub
    ''' <summary>
    ''' Initializes a service with a configuration object and a caller-managed HTTP client.
    ''' </summary>
    ''' <param name="Options">The Firebase configuration.</param>
    ''' <param name="HttpClient">The HTTP client to reuse. The service does not dispose it.</param>
    Public Sub New(Options As FirebaseOptions, HttpClient As HttpClient)
        Initialize(Options, HttpClient)
    End Sub
    ''' <summary>
    ''' Initializes or reinitializes the service with an internally managed HTTP client.
    ''' </summary>
    ''' <param name="Options">The Firebase configuration.</param>
    Public Sub Initialize(Options As FirebaseOptions)
        Dim ManagedHttpClient As New HttpClient()
        Try
            InitializeCore(Options, ManagedHttpClient, True)
        Catch
            ManagedHttpClient.Dispose()
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' Initializes or reinitializes the service with a caller-managed HTTP client.
    ''' </summary>
    ''' <param name="Options">The Firebase configuration.</param>
    ''' <param name="HttpClient">The HTTP client to reuse. The service does not dispose it.</param>
    Public Sub Initialize(Options As FirebaseOptions, HttpClient As HttpClient)
        InitializeCore(Options, HttpClient, False)
    End Sub
    ''' <summary>
    ''' Initializes or reinitializes the service from the required Firebase project values.
    ''' </summary>
    ''' <param name="ApiKey">The Firebase Web API key.</param>
    ''' <param name="ProjectId">The Firebase project identifier.</param>
    ''' <param name="StorageBucket">The Cloud Storage bucket name.</param>
    Public Sub Initialize(ApiKey As String, ProjectId As String, StorageBucket As String)
        Initialize(New FirebaseOptions(ApiKey, ProjectId, StorageBucket))
    End Sub
    Private Sub InitializeCore(Options As FirebaseOptions, HttpClient As HttpClient, OwnsHttpClient As Boolean)
        ArgumentNullException.ThrowIfNull(Options)
        ArgumentNullException.ThrowIfNull(HttpClient)
        Dim NewClient As New FirebaseClient(Options, HttpClient, OwnsHttpClient)
        Dim NewAuth As New FirebaseAuth(NewClient)
        Dim NewFirestore As New FirebaseFirestore(NewClient)
        Dim NewStorage As New FirebaseStorage(NewClient)
        NewClient.Auth = NewAuth
        SyncLock _SyncRoot
            If _Disposed Then
                NewClient.Dispose()
                Throw New ObjectDisposedException(NameOf(FirebaseService))
            End If
            Dim PreviousClient As FirebaseClient = _Client
            _Client = NewClient
            _Auth = NewAuth
            _Firestore = NewFirestore
            _Storage = NewStorage
            If PreviousClient IsNot Nothing Then PreviousClient.Dispose()
        End SyncLock
    End Sub
    Private Function GetRequiredClient() As FirebaseClient
        ObjectDisposedException.ThrowIf(_Disposed, Me)
        If _Client Is Nothing Then Throw New InvalidOperationException("FirebaseService has not been initialized. Call Initialize before accessing its modules.")
        Return _Client
    End Function
    ''' <summary>
    ''' Releases the internally managed HTTP client and clears the in-memory Firebase session.
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        SyncLock _SyncRoot
            If _Disposed Then Return
            _Disposed = True
            If _Auth IsNot Nothing Then _Auth.Logout()
            If _Client IsNot Nothing Then _Client.Dispose()
            _Client = Nothing
            _Auth = Nothing
            _Firestore = Nothing
            _Storage = Nothing
        End SyncLock
    End Sub
End Class
