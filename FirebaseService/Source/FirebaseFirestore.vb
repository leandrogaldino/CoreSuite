Imports System.Net.Http
Imports System.Text
Imports CoreSuite.Helpers
Imports System.Web.Script.Serialization
''' <summary>
''' Provides access to Firebase Firestore database operations.
''' </summary>
Public Class FirebaseFirestore
    Private ReadOnly _Client As FirebaseClient
    Private Const DocumentIDFieldName = "firestore_document_id"
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Retrieves all sub-collections of a document or root collections.
    ''' </summary>
    ''' <param name="DocumentPath">Optional document path.</param>
    ''' <returns>A dictionary containing collection identifiers.</returns>
    Public Async Function GetCollectionsAsync(Optional DocumentPath = Nothing) As Task(Of Dictionary(Of String, Object))
        Try
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{DocumentPath}:listCollectionIds"
            Using Request = _Client.CreateRequest(HttpMethod.Get, Url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    Dim JsonRaw = Await Response.Content.ReadAsStringAsync()
                    Return FirestoreJsonToMap(JsonRaw)
                Else
                    Return New Dictionary(Of String, Object)()
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Falha de rede ao conectar com Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Retrieves a document from a Firestore collection.
    ''' </summary>
    ''' <param name="Collection">Collection name.</param>
    ''' <param name="DocumentID">Document identifier.</param>
    ''' <returns>Document data as a dictionary.</returns>
    Public Async Function GetDocumentAsync(Collection As String, DocumentID As String) As Task(Of Dictionary(Of String, Object))
        Try
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}/{DocumentID}"
            Using Request = _Client.CreateRequest(HttpMethod.Get, Url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    Dim JsonRaw = Await Response.Content.ReadAsStringAsync()
                    Return FirestoreJsonToMap(JsonRaw)
                Else
                    Return New Dictionary(Of String, Object)()
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Falha de rede ao conectar com Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Retrieves all documents from a Firestore collection.
    ''' </summary>
    ''' <param name="Collection">Collection name.</param>
    ''' <returns>A list of documents represented as dictionaries.</returns>
    Public Async Function GetAllDocumentsAsync(Collection As String) As Task(Of List(Of Dictionary(Of String, Object)))
        Try
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}"
            Dim ResultList As New List(Of Dictionary(Of String, Object))
            Using Request = _Client.CreateRequest(HttpMethod.Get, Url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                Dim JsonRaw = Await Response.Content.ReadAsStringAsync()
                If Response.IsSuccessStatusCode Then
                    Dim Jss As New JavaScriptSerializer()
                    Dim Root = Jss.Deserialize(Of Dictionary(Of String, Object))(JsonRaw)
                    If Root IsNot Nothing AndAlso Root.ContainsKey("documents") Then
                        Dim Documents = TryCast(Root("documents"), IEnumerable)
                        For Each Document In Documents
                            Dim DocumentJson = Jss.Serialize(Document)
                            ResultList.Add(FirestoreJsonToMap(DocumentJson))
                        Next
                    End If
                    Return ResultList
                Else
                    Throw New Exception($"Erro ao listar coleção: {JsonRaw}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Falha de rede ao conectar com Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Executes a composite query using multiple filters.
    ''' </summary>
    ''' <param name="Collection">Target collection.</param>
    ''' <param name="Filters">List of query filters.</param>
    ''' <returns>Matching documents.</returns>
    Public Async Function QueryCompositeAsync(Collection As String, Filters As List(Of FirestoreFilter)) As Task(Of List(Of Dictionary(Of String, Object)))
        Try
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents:runQuery"
            Dim JsonFilterList As New List(Of String)
            For Each FirestoreFilter In Filters
                Dim JsonValue = If(FirestoreFilter.Operator = FirestoreOperator.InList OrElse FirestoreFilter.Operator = FirestoreOperator.NotInList,
                           """arrayValue"": { ""values"": [" & String.Join(",", DirectCast(FirestoreFilter.Value, IEnumerable).Cast(Of Object).Select(Function(x) "{" & FormatFirestoreValue(x) & "}")) & "] }",
                           FormatFirestoreValue(FirestoreFilter.Value))

                JsonFilterList.Add("{ ""fieldFilter"": { " &
                    """field"": { ""fieldPath"": """ & FirestoreFilter.Field & """ }," &
                    """op"": """ & GetOperatorString(FirestoreFilter.Operator) & """," &
                    """value"": { " & JsonValue & " }" &
                "} }")
            Next
            Dim QueryJson = "{" &
                """structuredQuery"": {" &
                    """from"": [{ ""collectionId"": """ & Collection & """ }]," &
                    """where"": {" &
                        """compositeFilter"": {" &
                            """op"": ""AND""," &
                            """filters"": [" & String.Join(",", JsonFilterList) & "]" &
                        "}" &
                    "}" &
                "}" &
            "}"
            Dim ResultList As New List(Of Dictionary(Of String, Object))
            Using Request = _Client.CreateRequest(HttpMethod.Post, Url)
                Request.Content = New StringContent(QueryJson, Encoding.UTF8, "application/json")
                Dim Resposnse = Await _Client.Http.SendAsync(Request)
                Dim JsonRaw = Await Resposnse.Content.ReadAsStringAsync()
                If Resposnse.IsSuccessStatusCode Then
                    Dim Jss As New JavaScriptSerializer()
                    Jss.MaxJsonLength = Int32.MaxValue
                    Dim Results = Jss.Deserialize(Of List(Of Dictionary(Of String, Object)))(JsonRaw)
                    If Results IsNot Nothing Then
                        For Each Item In Results
                            If Item IsNot Nothing AndAlso Item.ContainsKey("document") Then
                                Dim docJson = Jss.Serialize(Item("document"))
                                ResultList.Add(FirestoreJsonToMap(docJson))
                            End If
                        Next
                    End If
                    Return ResultList
                Else
                    Throw New Exception($"Erro na Query: {JsonRaw}")
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Creates or updates a Firestore document.
    ''' </summary>
    ''' <param name="Collection">Collection name.</param>
    ''' <param name="DocumentID">Document identifier or empty for auto-ID.</param>
    ''' <param name="Fields">Document fields.</param>
    ''' <returns>The document identifier.</returns>
    Public Async Function SaveDocumentAsync(Collection As String, DocumentID As String, Fields As Dictionary(Of String, Object)) As Task(Of String)
        Try
            Dim IsAutoId As Boolean = String.IsNullOrEmpty(DocumentID)
            Dim Url As String
            Dim Method As HttpMethod
            If IsAutoId Then
                Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}"
                Method = HttpMethod.Post
            Else
                Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}/{DocumentID}"
                Method = New HttpMethod("PATCH")
            End If
            Dim JsonPayload = MapToFirestoreJson(Fields)
            Using Request = _Client.CreateRequest(Method, Url)
                Request.Content = New StringContent(JsonPayload, Encoding.UTF8, "application/json")
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    If IsAutoId Then
                        Dim ResponseBody = Await Response.Content.ReadAsStringAsync()
                        Dim FullName = TextHelper.ExtractJsonValue(ResponseBody, "name")
                        Return FullName.Split("/"c).Last()
                    Else
                        Return DocumentID
                    End If
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Falha de rede ao conectar com Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
        Return String.Empty
    End Function


    ''' <summary>
    ''' Deletes a document from a Firestore collection.
    ''' </summary>
    ''' <param name="Collection">Collection name.</param>
    ''' <param name="DocumentID">Document identifier.</param>
    ''' <returns>True if deletion was successful.</returns>
    Public Async Function DeleteDocumentAsync(Collection As String, DocumentID As String) As Task(Of Boolean)
        Try
            Dim url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}/{DocumentID}"
            Using Request = _Client.CreateRequest(HttpMethod.Delete, url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    Return True
                Else
                    Dim ErrorJson As String = Await Response.Content.ReadAsStringAsync()
                    Throw New Exception($"Erro ao deletar documento: {ErrorJson}")
                End If
            End Using

        Catch ex As HttpRequestException
            Throw New Exception("Falha de rede ao conectar com Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Determines whether a document exists in a collection.
    ''' </summary>
    ''' <param name="Collection">Collection name.</param>
    ''' <param name="DocumentID">Document identifier.</param>
    ''' <returns>True if the document exists.</returns>
    Public Async Function DocumentExistsAsync(Collection As String, DocumentID As String) As Task(Of Boolean)
        Try
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}/{DocumentID}"
            Using Request = _Client.CreateRequest(HttpMethod.Get, Url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    Return True
                ElseIf Response.StatusCode = Net.HttpStatusCode.NotFound Then
                    Return False
                Else
                    Dim [Error] = Await Response.Content.ReadAsStringAsync()
                    Throw New Exception($"Erro ao verificar existência: {[Error]}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Falha de rede ao conectar com Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function FormatFirestoreValue(Value As Object) As String
        If TypeOf Value Is Boolean Then
            Return """booleanValue"": " & Value.ToString().ToLower()
        ElseIf TypeOf Value Is DateTime Then
            Return """timestampValue"": """ & DirectCast(Value, DateTime).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ") & """"
        ElseIf IsNumeric(Value) Then
            Return """doubleValue"": " & Convert.ToDouble(Value).ToString(Globalization.CultureInfo.InvariantCulture)
        Else
            Return """stringValue"": """ & Value.ToString() & """"
        End If
    End Function
    Private Function MapToFirestoreJson(Dictionary As Dictionary(Of String, Object)) As String
        Dim Builder As New StringBuilder()
        Builder.Append("{""fields"": {")
        Dim Entries As New List(Of String)()
        For Each Kvp In Dictionary
            If Kvp.Key = DocumentIDFieldName Then Continue For
            Dim ValuePart As String = ""
            If TypeOf Kvp.Value Is String Then
                ValuePart = "{""stringValue"": """ & Kvp.Value.ToString() & """}"
            ElseIf TypeOf Kvp.Value Is Integer OrElse TypeOf Kvp.Value Is Long Then
                ValuePart = "{""integerValue"": " & Kvp.Value.ToString() & "}"
            ElseIf TypeOf Kvp.Value Is Boolean Then
                ValuePart = "{""booleanValue"": " & Kvp.Value.ToString().ToLower() & "}"
            ElseIf TypeOf Kvp.Value Is Double OrElse TypeOf Kvp.Value Is Decimal Then
                Dim doubleVal As Double = Convert.ToDouble(Kvp.Value)
                ValuePart = "{""doubleValue"": " & doubleVal.ToString(System.Globalization.CultureInfo.InvariantCulture) & "}"
            ElseIf TypeOf Kvp.Value Is DateTime Then
                Dim [Date] As DateTime = DirectCast(Kvp.Value, DateTime)
                Dim TimestampIso = [Date].ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                ValuePart = "{""timestampValue"": """ & TimestampIso & """}"
            End If
            If Not String.IsNullOrEmpty(ValuePart) Then
                Entries.Add($"""{Kvp.Key}"": {ValuePart}")
            End If
        Next
        Builder.Append(String.Join(",", Entries))
        Builder.Append("}}")
        Return Builder.ToString()
    End Function
    Private Function FirestoreJsonToMap(JsonRaw As String) As Dictionary(Of String, Object)
        Dim Resultado As New Dictionary(Of String, Object)
        Try
            Dim Jss As New Web.Script.Serialization.JavaScriptSerializer()
            Jss.MaxJsonLength = Int32.MaxValue
            Dim Root = Jss.Deserialize(Of Dictionary(Of String, Object))(JsonRaw)
            If Root IsNot Nothing Then
                If Root.ContainsKey("name") Then
                    Dim FullPath As String = Root("name").ToString()
                    Dim DocumentID As String = FullPath.Split("/"c).Last()
                    Resultado.Add(DocumentIDFieldName, DocumentID)
                End If
                If Root.ContainsKey("fields") Then
                    Dim Fields = DirectCast(Root("fields"), Dictionary(Of String, Object))
                    For Each Field In Fields
                        If Field.Key = DocumentIDFieldName Then Continue For
                        Dim TypeAndValue = DirectCast(Field.Value, Dictionary(Of String, Object))
                        For Each Kvp In TypeAndValue
                            Dim RealValue As Object = Kvp.Value
                            Select Case Kvp.Key
                                Case "integerValue"
                                    Resultado.Add(Field.Key, Convert.ToInt64(RealValue))
                                Case "doubleValue"
                                    Resultado.Add(Field.Key, Convert.ToDouble(RealValue, System.Globalization.CultureInfo.InvariantCulture))
                                Case "booleanValue"
                                    Resultado.Add(Field.Key, Convert.ToBoolean(RealValue))
                                Case "stringValue"
                                    Resultado.Add(Field.Key, Convert.ToString(RealValue))
                                Case "timestampValue"
                                    Resultado.Add(Field.Key, Convert.ToDateTime(RealValue))
                                Case Else
                                    Resultado.Add(Field.Key, RealValue.ToString())
                            End Select
                            Exit For
                        Next
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception($"Erro ao converter JSON do Firestore para Dictionary.{Environment.NewLine}JSON: {JsonRaw}", ex)
        End Try
        Return Resultado
    End Function
    Private Function GetOperatorString([Operator] As FirestoreOperator) As String
        Select Case [Operator]
            Case FirestoreOperator.Equal : Return "EQUAL"
            Case FirestoreOperator.NotEqual : Return "NOT_EQUAL"
            Case FirestoreOperator.LessThan : Return "LESS_THAN"
            Case FirestoreOperator.LessThanOrEqual : Return "LESS_THAN_OR_EQUAL"
            Case FirestoreOperator.GreaterThan : Return "GREATER_THAN"
            Case FirestoreOperator.GreaterThanOrEqual : Return "GREATER_THAN_OR_EQUAL"
            Case FirestoreOperator.ArrayContains : Return "ARRAY_CONTAINS"
            Case FirestoreOperator.ArrayContainsAny : Return "ARRAY_CONTAINS_ANY"
            Case FirestoreOperator.InList : Return "IN"
            Case FirestoreOperator.NotInList : Return "NOT_IN"
            Case Else : Return "EQUAL"
        End Select
    End Function
End Class