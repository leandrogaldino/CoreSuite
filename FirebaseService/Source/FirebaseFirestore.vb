Imports System.Collections
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading
''' <summary>
''' Provides authenticated Cloud Firestore document, collection and query operations.
''' </summary>
Public NotInheritable Class FirebaseFirestore
    Private Const PageSize As Integer = 1000
    Private ReadOnly _Client As FirebaseClient
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Lists all collection identifiers at the database root or directly beneath a document.
    ''' </summary>
    ''' <param name="DocumentPath">An optional relative document path. Omit it to list root collections.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns>All collection identifiers across every response page.</returns>
    Public Async Function GetCollectionsAsync(Optional DocumentPath As String = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of List(Of String))
        Dim NormalizedDocumentPath As String = Nothing
        If Not String.IsNullOrWhiteSpace(DocumentPath) Then NormalizedDocumentPath = FirebasePathHelper.NormalizeDocumentPath(DocumentPath, NameOf(DocumentPath))
        Dim ParentUrl As String = GetDocumentsBaseUrl()
        If NormalizedDocumentPath IsNot Nothing Then ParentUrl &= "/" & FirebasePathHelper.EncodeFirestorePath(NormalizedDocumentPath)
        Dim Url As String = ParentUrl & ":listCollectionIds"
        Dim CollectionIds As New List(Of String)()
        Dim PageToken As String = Nothing
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Do
                    Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                    Dim Payload As New JsonObject()
                    Payload("pageSize") = JsonValue.Create(PageSize)
                    If Not String.IsNullOrWhiteSpace(PageToken) Then Payload("pageToken") = JsonValue.Create(PageToken)
                    Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Post, Url)
                        Request.Content = New StringContent(Payload.ToJsonString(), Encoding.UTF8, "application/json")
                        Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                            Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                            Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                            Using Document As JsonDocument = ParseJson(ResponseBody, "collection list")
                                Dim CollectionIdsElement As JsonElement
                                If Document.RootElement.TryGetProperty("collectionIds", CollectionIdsElement) Then
                                    If CollectionIdsElement.ValueKind <> JsonValueKind.Array Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, "Cloud Firestore returned an invalid collectionIds value.")
                                    For Each CollectionIdElement As JsonElement In CollectionIdsElement.EnumerateArray()
                                        If CollectionIdElement.ValueKind <> JsonValueKind.String Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, "Cloud Firestore returned a non-string collection identifier.")
                                        CollectionIds.Add(CollectionIdElement.GetString())
                                    Next CollectionIdElement
                                End If
                                PageToken = ReadOptionalString(Document.RootElement, "nextPageToken")
                            End Using
                        End Using
                    End Using
                Loop While Not String.IsNullOrWhiteSpace(PageToken)
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
        Return CollectionIds
    End Function
    ''' <summary>
    ''' Retrieves a document by collection path and document identifier.
    ''' </summary>
    ''' <param name="CollectionPath">The relative collection path.</param>
    ''' <param name="DocumentId">The document identifier.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns>The document, or <see langword="Nothing"/> when Firestore returns HTTP 404.</returns>
    Public Async Function GetDocumentAsync(CollectionPath As String, DocumentId As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of FirestoreDocument)
        Dim DocumentPath As String = BuildDocumentPath(CollectionPath, DocumentId)
        Dim Url As String = GetDocumentsBaseUrl() & "/" & FirebasePathHelper.EncodeFirestorePath(DocumentPath)
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Get, Url)
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                        If Response.StatusCode = HttpStatusCode.NotFound Then Return Nothing
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                        Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                        Return ParseDocument(ResponseBody)
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Retrieves every document from a collection, following all Firestore response pages.
    ''' </summary>
    ''' <param name="CollectionPath">The relative collection path.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns>All documents returned from the collection.</returns>
    Public Async Function GetAllDocumentsAsync(CollectionPath As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of List(Of FirestoreDocument))
        Dim NormalizedCollectionPath As String = FirebasePathHelper.NormalizeCollectionPath(CollectionPath, NameOf(CollectionPath))
        Dim CollectionUrl As String = GetDocumentsBaseUrl() & "/" & FirebasePathHelper.EncodeFirestorePath(NormalizedCollectionPath)
        Dim Documents As New List(Of FirestoreDocument)()
        Dim PageToken As String = Nothing
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Do
                    Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                    Dim Url As String = CollectionUrl & $"?pageSize={PageSize}"
                    If Not String.IsNullOrWhiteSpace(PageToken) Then Url &= "&pageToken=" & Uri.EscapeDataString(PageToken)
                    Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Get, Url)
                        Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                            Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                            Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                            Using Document As JsonDocument = ParseJson(ResponseBody, "document list")
                                Dim DocumentsElement As JsonElement
                                If Document.RootElement.TryGetProperty("documents", DocumentsElement) Then
                                    If DocumentsElement.ValueKind <> JsonValueKind.Array Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, "Cloud Firestore returned an invalid documents value.")
                                    For Each DocumentElement As JsonElement In DocumentsElement.EnumerateArray()
                                        Documents.Add(ParseDocument(DocumentElement))
                                    Next DocumentElement
                                End If
                                PageToken = ReadOptionalString(Document.RootElement, "nextPageToken")
                            End Using
                        End Using
                    End Using
                Loop While Not String.IsNullOrWhiteSpace(PageToken)
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
        Return Documents
    End Function
    ''' <summary>
    ''' Executes a structured Firestore query whose filters are combined with logical <c>AND</c>.
    ''' </summary>
    ''' <param name="CollectionPath">The relative collection path to query.</param>
    ''' <param name="Filters">One or more field filters.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns>The matching documents.</returns>
    Public Async Function QueryCompositeAsync(CollectionPath As String, Filters As IEnumerable(Of FirestoreFilter), Optional CancellationToken As CancellationToken = Nothing) As Task(Of List(Of FirestoreDocument))
        Dim NormalizedCollectionPath As String = FirebasePathHelper.NormalizeCollectionPath(CollectionPath, NameOf(CollectionPath))
        ArgumentNullException.ThrowIfNull(Filters)
        Dim FilterList As List(Of FirestoreFilter) = Filters.ToList()
        If FilterList.Count = 0 Then Throw New ArgumentException("At least one Firestore filter is required.", NameOf(Filters))
        For Each Filter As FirestoreFilter In FilterList
            ValidateFilter(Filter)
        Next Filter
        Dim Segments As String() = NormalizedCollectionPath.Split("/"c)
        Dim CollectionId As String = Segments.Last()
        Dim ParentDocumentPath As String = If(Segments.Length = 1, Nothing, String.Join("/", Segments.Take(Segments.Length - 1)))
        Dim Url As String = GetDocumentsBaseUrl()
        If ParentDocumentPath IsNot Nothing Then Url &= "/" & FirebasePathHelper.EncodeFirestorePath(ParentDocumentPath)
        Url &= ":runQuery"
        Dim Payload As JsonObject = CreateQueryPayload(CollectionId, FilterList)
        Dim Results As New List(Of FirestoreDocument)()
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Post, Url)
                    Request.Content = New StringContent(Payload.ToJsonString(), Encoding.UTF8, "application/json")
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                        Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                        Using Document As JsonDocument = ParseJson(ResponseBody, "query result")
                            If Document.RootElement.ValueKind <> JsonValueKind.Array Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, "Cloud Firestore returned an invalid query result.")
                            For Each ResultElement As JsonElement In Document.RootElement.EnumerateArray()
                                Dim DocumentElement As JsonElement
                                If ResultElement.ValueKind = JsonValueKind.Object AndAlso ResultElement.TryGetProperty("document", DocumentElement) Then Results.Add(ParseDocument(DocumentElement))
                            Next ResultElement
                        End Using
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
        Return Results
    End Function
    ''' <summary>
    ''' Creates a document with an automatic identifier or replaces a document with a caller-supplied identifier.
    ''' </summary>
    ''' <param name="CollectionPath">The relative collection path.</param>
    ''' <param name="DocumentId">The document identifier, or <see langword="Nothing"/> or an empty string to request an automatic identifier.</param>
    ''' <param name="Fields">The complete set of document fields.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns>The identifier returned by Firestore.</returns>
    Public Async Function SaveDocumentAsync(CollectionPath As String, DocumentId As String, Fields As IDictionary(Of String, Object), Optional CancellationToken As CancellationToken = Nothing) As Task(Of String)
        Dim NormalizedCollectionPath As String = FirebasePathHelper.NormalizeCollectionPath(CollectionPath, NameOf(CollectionPath))
        Dim UsesAutomaticId As Boolean = String.IsNullOrWhiteSpace(DocumentId)
        Dim Url As String
        Dim Method As HttpMethod
        If UsesAutomaticId Then
            Url = GetDocumentsBaseUrl() & "/" & FirebasePathHelper.EncodeFirestorePath(NormalizedCollectionPath)
            Method = HttpMethod.Post
        Else
            Dim NormalizedDocumentId As String = FirebasePathHelper.NormalizeDocumentId(DocumentId, NameOf(DocumentId))
            Url = GetDocumentsBaseUrl() & "/" & FirebasePathHelper.EncodeFirestorePath(NormalizedCollectionPath & "/" & NormalizedDocumentId)
            Method = HttpMethod.Patch
        End If
        Dim Payload As JsonObject = FirestoreValueConverter.SerializeDocument(Fields)
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(Method, Url)
                    Request.Content = New StringContent(Payload.ToJsonString(), Encoding.UTF8, "application/json")
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                        Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                        Return ParseDocument(ResponseBody).Id
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Deletes a Firestore document.
    ''' </summary>
    ''' <param name="CollectionPath">The relative collection path.</param>
    ''' <param name="DocumentId">The document identifier.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns><see langword="True"/> when the document was deleted; <see langword="False"/> when it did not exist.</returns>
    Public Async Function DeleteDocumentAsync(CollectionPath As String, DocumentId As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        Dim DocumentPath As String = BuildDocumentPath(CollectionPath, DocumentId)
        Dim Url As String = GetDocumentsBaseUrl() & "/" & FirebasePathHelper.EncodeFirestorePath(DocumentPath)
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Delete, Url)
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                        If Response.StatusCode = HttpStatusCode.NotFound Then Return False
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                        Return True
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Determines whether a Firestore document exists.
    ''' </summary>
    ''' <param name="CollectionPath">The relative collection path.</param>
    ''' <param name="DocumentId">The document identifier.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns><see langword="True"/> when the document exists; otherwise, <see langword="False"/>.</returns>
    Public Async Function DocumentExistsAsync(CollectionPath As String, DocumentId As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        Dim DocumentPath As String = BuildDocumentPath(CollectionPath, DocumentId)
        Dim Url As String = GetDocumentsBaseUrl() & "/" & FirebasePathHelper.EncodeFirestorePath(DocumentPath)
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(False, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Get, Url)
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Firestore, HttpCompletionOption.ResponseHeadersRead, OperationSource.Token).ConfigureAwait(False)
                        If Response.StatusCode = HttpStatusCode.NotFound Then Return False
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Firestore, OperationSource.Token).ConfigureAwait(False)
                        Return True
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Firestore, ex)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Creates a typed Firestore reference value for a relative document path in the configured database.
    ''' </summary>
    ''' <param name="DocumentPath">The relative document path.</param>
    ''' <returns>A reference that can be stored in a Firestore document or query filter.</returns>
    Public Function CreateDocumentReference(DocumentPath As String) As FirestoreDocumentReference
        Dim NormalizedDocumentPath As String = FirebasePathHelper.NormalizeDocumentPath(DocumentPath, NameOf(DocumentPath))
        Return New FirestoreDocumentReference(_Client.FirestoreDocumentRoot & "/" & NormalizedDocumentPath)
    End Function
    Private Function GetDocumentsBaseUrl() As String
        Return $"https://firestore.googleapis.com/v1/projects/{Uri.EscapeDataString(_Client.Options.ProjectId)}/databases/{Uri.EscapeDataString(_Client.Options.DatabaseId)}/documents"
    End Function
    Private Shared Function BuildDocumentPath(CollectionPath As String, DocumentId As String) As String
        Dim NormalizedCollectionPath As String = FirebasePathHelper.NormalizeCollectionPath(CollectionPath, NameOf(CollectionPath))
        Dim NormalizedDocumentId As String = FirebasePathHelper.NormalizeDocumentId(DocumentId, NameOf(DocumentId))
        Return NormalizedCollectionPath & "/" & NormalizedDocumentId
    End Function
    Private Shared Function CreateQueryPayload(CollectionId As String, Filters As List(Of FirestoreFilter)) As JsonObject
        Dim FromValue As New JsonObject()
        FromValue("collectionId") = JsonValue.Create(CollectionId)
        Dim FromValues As New JsonArray From {
            FromValue
        }
        Dim StructuredQuery As New JsonObject()
        StructuredQuery("from") = FromValues
        If Filters.Count = 1 Then
            StructuredQuery("where") = CreateFieldFilter(Filters(0))
        Else
            Dim FilterValues As New JsonArray()
            For Each Filter As FirestoreFilter In Filters
                FilterValues.Add(CreateFieldFilter(Filter))
            Next Filter
            Dim CompositeFilter As New JsonObject()
            CompositeFilter("op") = JsonValue.Create("AND")
            CompositeFilter("filters") = FilterValues
            Dim WhereValue As New JsonObject()
            WhereValue("compositeFilter") = CompositeFilter
            StructuredQuery("where") = WhereValue
        End If
        Dim Payload As New JsonObject()
        Payload("structuredQuery") = StructuredQuery
        Return Payload
    End Function
    Private Shared Function CreateFieldFilter(Filter As FirestoreFilter) As JsonObject
        Dim FieldReference As New JsonObject()
        FieldReference("fieldPath") = JsonValue.Create(Filter.Field)
        Dim FieldFilter As New JsonObject()
        FieldFilter("field") = FieldReference
        FieldFilter("op") = JsonValue.Create(GetOperatorString(Filter.Operator))
        FieldFilter("value") = FirestoreValueConverter.SerializeValue(Filter.Value)
        Dim FilterValue As New JsonObject()
        FilterValue("fieldFilter") = FieldFilter
        Return FilterValue
    End Function
    Private Shared Sub ValidateFilter(Filter As FirestoreFilter)
        If Filter Is Nothing Then Throw New ArgumentException("The filter collection cannot contain Nothing.", NameOf(Filter))
        If String.IsNullOrWhiteSpace(Filter.Field) Then Throw New ArgumentException("Firestore filter field paths cannot be empty.", NameOf(Filter))
        If Not [Enum].IsDefined(GetType(FirestoreOperator), Filter.Operator) Then Throw New ArgumentOutOfRangeException(NameOf(Filter), "The Firestore operator is invalid.")
        If Filter.Operator = FirestoreOperator.ArrayContainsAny OrElse Filter.Operator = FirestoreOperator.InList OrElse Filter.Operator = FirestoreOperator.NotInList Then
            If Filter.Value Is Nothing OrElse TypeOf Filter.Value Is String OrElse TypeOf Filter.Value IsNot IEnumerable OrElse TypeOf Filter.Value Is IDictionary Then Throw New ArgumentException($"The {Filter.Operator} operator requires an enumerable value.", NameOf(Filter))
        End If
    End Sub
    Private Shared Function GetOperatorString(OperatorValue As FirestoreOperator) As String
        Select Case OperatorValue
            Case FirestoreOperator.Equal
                Return "EQUAL"
            Case FirestoreOperator.NotEqual
                Return "NOT_EQUAL"
            Case FirestoreOperator.LessThan
                Return "LESS_THAN"
            Case FirestoreOperator.LessThanOrEqual
                Return "LESS_THAN_OR_EQUAL"
            Case FirestoreOperator.GreaterThan
                Return "GREATER_THAN"
            Case FirestoreOperator.GreaterThanOrEqual
                Return "GREATER_THAN_OR_EQUAL"
            Case FirestoreOperator.ArrayContains
                Return "ARRAY_CONTAINS"
            Case FirestoreOperator.ArrayContainsAny
                Return "ARRAY_CONTAINS_ANY"
            Case FirestoreOperator.InList
                Return "IN"
            Case FirestoreOperator.NotInList
                Return "NOT_IN"
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(OperatorValue))
        End Select
    End Function
    Private Shared Function ParseDocument(ResponseBody As String) As FirestoreDocument
        Using Document As JsonDocument = ParseJson(ResponseBody, "document")
            Return ParseDocument(Document.RootElement)
        End Using
    End Function
    Private Shared Function ParseDocument(DocumentElement As JsonElement) As FirestoreDocument
        Try
            Return FirestoreValueConverter.DeserializeDocument(DocumentElement)
        Catch ex As JsonException
            Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, "Cloud Firestore returned an invalid document.", ex)
        Catch ex As FormatException
            Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, "Cloud Firestore returned a document containing an invalid value.", ex)
        End Try
    End Function
    Private Shared Function ParseJson(ResponseBody As String, Description As String) As JsonDocument
        Try
            Return JsonDocument.Parse(ResponseBody)
        Catch ex As JsonException
            Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, $"Cloud Firestore returned malformed JSON for the {Description}.", ex)
        End Try
    End Function
    Private Shared Function ReadOptionalString(Element As JsonElement, PropertyName As String) As String
        Dim PropertyElement As JsonElement
        If Not Element.TryGetProperty(PropertyName, PropertyElement) Then Return Nothing
        If PropertyElement.ValueKind <> JsonValueKind.String Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Firestore, $"Cloud Firestore returned an invalid {PropertyName} value.")
        Return PropertyElement.GetString()
    End Function
End Class
