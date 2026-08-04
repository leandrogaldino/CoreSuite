Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading
Friend NotInheritable Class FirebaseClient
    Implements IDisposable
    Private ReadOnly _OwnsHttpClient As Boolean
    Private _Disposed As Boolean
    Friend ReadOnly Property Options As FirebaseOptions
    Friend ReadOnly Property Http As HttpClient
    Friend Property Auth As FirebaseAuth
    Friend Property Token As String
    Friend ReadOnly Property FirestoreDocumentRoot As String
        Get
            Return $"projects/{Options.ProjectId}/databases/{Options.DatabaseId}/documents"
        End Get
    End Property
    Friend Sub New(Options As FirebaseOptions, HttpClient As HttpClient, OwnsHttpClient As Boolean)
        Me.Options = Options.CreateValidatedCopy()
        ArgumentNullException.ThrowIfNull(HttpClient)
        Me.Http = HttpClient
        _OwnsHttpClient = OwnsHttpClient
        If OwnsHttpClient Then Me.Http.Timeout = Timeout.InfiniteTimeSpan
    End Sub
    Friend Function CreateRequest(Method As HttpMethod, Url As String) As HttpRequestMessage
        ThrowIfDisposed()
        Dim Request As New HttpRequestMessage(Method, Url)
        If Not String.IsNullOrWhiteSpace(Token) Then Request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", Token)
        Return Request
    End Function
    Friend Function CreateOperationCancellationSource(IsTransfer As Boolean, CancellationToken As CancellationToken) As CancellationTokenSource
        ThrowIfDisposed()
        Dim Source As CancellationTokenSource = CancellationTokenSource.CreateLinkedTokenSource(CancellationToken)
        Dim OperationTimeout As TimeSpan = If(IsTransfer, Options.TransferTimeout, Options.RequestTimeout)
        If OperationTimeout <> Timeout.InfiniteTimeSpan Then Source.CancelAfter(OperationTimeout)
        Return Source
    End Function
    Friend Async Function SendAsync(Request As HttpRequestMessage, ServiceArea As FirebaseServiceArea, CompletionOption As HttpCompletionOption, CancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
        ThrowIfDisposed()
        Try
            Return Await Http.SendAsync(Request, CompletionOption, CancellationToken).ConfigureAwait(False)
        Catch ex As HttpRequestException
            Throw FirebaseException.CreateNetwork(ServiceArea, ex)
        End Try
    End Function
    Friend Shared Async Function EnsureSuccessAsync(Response As HttpResponseMessage, ServiceArea As FirebaseServiceArea, CancellationToken As CancellationToken) As Task
        If Response.IsSuccessStatusCode Then Return
        Dim RequestException As FirebaseException = Await FirebaseException.FromResponseAsync(ServiceArea, Response, CancellationToken).ConfigureAwait(False)
        Throw RequestException
    End Function
    Private Sub ThrowIfDisposed()
        ObjectDisposedException.ThrowIf(_Disposed, Me)
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        If _Disposed Then Return
        _Disposed = True
        Token = Nothing
        If _OwnsHttpClient Then Http.Dispose()
    End Sub
End Class
