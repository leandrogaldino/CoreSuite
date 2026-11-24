Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
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
        If HttpClient Is Nothing Then Throw New ArgumentNullException("HttpClient")
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
    Friend Shared Function CreateJsonSettings() As JsonSerializerSettings
        Return New JsonSerializerSettings With {
            .DateParseHandling = DateParseHandling.None,
            .TypeNameHandling = TypeNameHandling.None,
            .MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            .MaxDepth = 64
        }
    End Function
    Friend Shared Function SerializeJson(Value As Object) As String
        Using TextWriter As New StringWriter(Global.System.Globalization.CultureInfo.InvariantCulture)
            Using Writer As New JsonTextWriter(TextWriter)
                JsonSerializer.Create(CreateJsonSettings()).Serialize(Writer, Value)
                Writer.Flush()
                Return TextWriter.ToString()
            End Using
        End Using
    End Function
    Friend Shared Function ParseJsonToken(Content As String) As JToken
        Using TextReader As New StringReader(Content)
            Using Reader As New JsonTextReader(TextReader) With {.DateParseHandling = DateParseHandling.None, .MaxDepth = 64}
                Dim Result As JToken = JToken.ReadFrom(Reader)
                If Reader.Read() Then Throw New JsonReaderException("Unexpected content after the JSON value.")
                Return Result
            End Using
        End Using
    End Function
    Friend Shared Function DeserializeJson(Of T)(Content As String) As T
        Dim Token As JToken = ParseJsonToken(Content)
        If Token.Type = JTokenType.Null Then Return Nothing
        Dim ObjectValue As JObject = TryCast(Token, JObject)
        If ObjectValue Is Nothing Then Throw New JsonSerializationException("The JSON response must be an object.")
        Dim Serializer As JsonSerializer = JsonSerializer.Create(CreateJsonSettings())
        Dim Contract As JsonObjectContract = DirectCast(Serializer.ContractResolver.ResolveContract(GetType(T)), JsonObjectContract)
        For Each PropertyValue As JProperty In ObjectValue.Properties().ToList()
            Dim PropertyContract As JsonProperty = Contract.Properties.GetProperty(PropertyValue.Name, StringComparison.Ordinal)
            If PropertyContract Is Nothing Then
                PropertyValue.Remove()
            ElseIf PropertyContract.PropertyType Is GetType(String) AndAlso PropertyValue.Value.Type <> JTokenType.String AndAlso PropertyValue.Value.Type <> JTokenType.Null Then
                Throw New JsonSerializationException($"The JSON property '{PropertyValue.Name}' must be a string.")
            End If
        Next
        Return ObjectValue.ToObject(Of T)(Serializer)
    End Function
    Friend Shared Function TryGetJsonProperty(Element As JToken, PropertyName As String, ByRef Value As JToken) As Boolean
        Dim ObjectValue As JObject = TryCast(Element, JObject)
        If ObjectValue Is Nothing Then Throw New InvalidOperationException("The JSON value is not an object.")
        Return ObjectValue.TryGetValue(PropertyName, StringComparison.Ordinal, Value)
    End Function
    Friend Shared Function ReadJsonString(Value As JToken) As String
        If Value.Type = JTokenType.Null Then Return Nothing
        If Value.Type <> JTokenType.String Then Throw New InvalidOperationException("The JSON value is not a string.")
        Return CType(Value, String)
    End Function
    Friend Shared Function ReadJsonBoolean(Value As JToken) As Boolean
        If Value.Type <> JTokenType.Boolean Then Throw New InvalidOperationException("The JSON value is not a Boolean.")
        Return CType(Value, Boolean)
    End Function
    Friend Shared Function IsJsonNumber(Value As JToken) As Boolean
        Return Value.Type = JTokenType.Integer OrElse Value.Type = JTokenType.Float
    End Function
    Friend Shared Function ReadJsonDouble(Value As JToken) As Double
        If Not IsJsonNumber(Value) Then Throw New InvalidOperationException("The JSON value is not a number.")
        Return CType(Value, Double)
    End Function
    Private Sub ThrowIfDisposed()
        If _Disposed Then Throw New ObjectDisposedException(Me.GetType().FullName)
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        If _Disposed Then Return
        _Disposed = True
        Token = Nothing
        If _OwnsHttpClient Then Http.Dispose()
    End Sub
End Class
