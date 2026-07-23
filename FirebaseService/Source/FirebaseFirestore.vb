Imports System.Net.Http
Imports System.Text
Imports System.Text.Json
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
            Await _Client.Auth.EnsureValidTokenAsync()
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
            Throw New Exception("Network failure while connecting to Firebase.", ex)
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
            Await _Client.Auth.EnsureValidTokenAsync()
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
            Throw New Exception("Network failure while connecting to Firebase.", ex)
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
            Await _Client.Auth.EnsureValidTokenAsync()
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}"
            Dim ResultList As New List(Of Dictionary(Of String, Object))
            Using Request = _Client.CreateRequest(HttpMethod.Get, Url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                Dim JsonRaw = Await Response.Content.ReadAsStringAsync()
                If Response.IsSuccessStatusCode Then
                    Dim Root As JsonElement = JsonSerializer.Deserialize(Of JsonElement)(JsonRaw)
                    If Root.TryGetProperty("documents", Nothing) Then
                        For Each Document As JsonElement In Root.GetProperty("documents").EnumerateArray()
                            Dim DocumentJson = Document.GetRawText()
                            ResultList.Add(FirestoreJsonToMap(DocumentJson))
                        Next Document
                    End If
                    Return ResultList
                Else
                    Throw New Exception($"Error listing collection: {JsonRaw}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Network failure while connecting to Firebase.", ex)
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
            Await _Client.Auth.EnsureValidTokenAsync()
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
            Next FirestoreFilter
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
                Dim Response = Await _Client.Http.SendAsync(Request)
                Dim JsonRaw = Await Response.Content.ReadAsStringAsync()
                If Response.IsSuccessStatusCode Then
                    Dim Results As JsonElement = JsonSerializer.Deserialize(Of JsonElement)(JsonRaw)
                    If Results.ValueKind = JsonValueKind.Array Then
                        For Each Item As JsonElement In Results.EnumerateArray()
                            Dim DocumentProperty As JsonElement
                            If Item.TryGetProperty("document", DocumentProperty) Then
                                Dim DocumentJson = DocumentProperty.GetRawText()
                                ResultList.Add(FirestoreJsonToMap(DocumentJson))
                            End If
                        Next Item
                    End If
                    Return ResultList
                Else
                    Throw New Exception($"Error in query: {JsonRaw}")
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
            Await _Client.Auth.EnsureValidTokenAsync()
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
            Throw New Exception("Network failure while connecting to Firebase.", ex)
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
            Await _Client.Auth.EnsureValidTokenAsync()
            Dim url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}/{DocumentID}"
            Using Request = _Client.CreateRequest(HttpMethod.Delete, url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    Return True
                Else
                    Dim ErrorJson As String = Await Response.Content.ReadAsStringAsync()
                    Throw New Exception($"Error deleting document: {ErrorJson}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Network failure while connecting to Firebase.", ex)
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
            Await _Client.Auth.EnsureValidTokenAsync()
            Dim Url = $"https://firestore.googleapis.com/v1/projects/{_Client.ProjectID}/databases/(default)/documents/{Collection}/{DocumentID}"
            Using Request = _Client.CreateRequest(HttpMethod.Get, Url)
                Dim Response = Await _Client.Http.SendAsync(Request)
                If Response.IsSuccessStatusCode Then
                    Return True
                ElseIf Response.StatusCode = Net.HttpStatusCode.NotFound Then
                    Return False
                Else
                    Dim [Error] = Await Response.Content.ReadAsStringAsync()
                    Throw New Exception($"Error checking existence: {[Error]}")
                End If
            End Using
        Catch ex As HttpRequestException
            Throw New Exception("Network failure while connecting to Firebase.", ex)
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Converts a value into its Firestore-formatted JSON field representation,
    ''' including the appropriate Firestore value type wrapper.
    ''' </summary>
    ''' <param name="Value">The value to format.</param>
    ''' <returns>A JSON string fragment representing the value in Firestore's field format.</returns>
    Private Shared Function FormatFirestoreValue(Value As Object) As String
        If Value Is Nothing Then
            Return """nullValue"": null"
        End If
        If TypeOf Value Is String Then
            Return """stringValue"": """ & EscapeJson(Value.ToString()) & """"
        End If
        If TypeOf Value Is Boolean Then
            Return """booleanValue"": " & Value.ToString().ToLower()
        End If
        If TypeOf Value Is Integer OrElse TypeOf Value Is Long Then
            Return """integerValue"": """ & Value.ToString() & """"
        End If
        If TypeOf Value Is Double OrElse TypeOf Value Is Decimal OrElse TypeOf Value Is Single Then
            Dim doubleVal As Double = Convert.ToDouble(Value)
            Return """doubleValue"": " & doubleVal.ToString(Globalization.CultureInfo.InvariantCulture)
        End If
        If TypeOf Value Is DateTime Then
            Dim Dt As DateTime = DirectCast(Value, DateTime)
            Dim Iso = Dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            Return """timestampValue"": """ & Iso & """"
        End If
        If TypeOf Value Is Dictionary(Of String, Object) Then
            Dim Dict = DirectCast(Value, Dictionary(Of String, Object))
            Dim Entries As New List(Of String)
            For Each Kvp In Dict
                Entries.Add($"""{Kvp.Key}"": {{ {FormatFirestoreValue(Kvp.Value)} }}")
            Next Kvp
            Return """mapValue"": { ""fields"": { " & String.Join(",", Entries) & " } }"
        End If
        If TypeOf Value Is IEnumerable AndAlso TypeOf Value IsNot String Then
            Dim Values As New List(Of String)
            For Each Item In DirectCast(Value, IEnumerable)
                Values.Add("{ " & FormatFirestoreValue(Item) & " }")
            Next Item
            Return """arrayValue"": { ""values"": [" & String.Join(",", Values) & "] }"
        End If
        Return """stringValue"": """ & EscapeJson(Value.ToString()) & """"
    End Function
    ''' <summary>
    ''' Converts a dictionary of field values into a complete Firestore document JSON payload.
    ''' </summary>
    ''' <param name="Dictionary">The dictionary containing the document fields and their values.</param>
    ''' <returns>A JSON string representing the Firestore document payload.</returns>
    Private Shared Function MapToFirestoreJson(Dictionary As Dictionary(Of String, Object)) As String
        Dim Builder As New StringBuilder()
        Builder.Append("{""fields"": {")
        Dim Entries As New List(Of String)
        For Each Kvp In Dictionary
            If Kvp.Key = DocumentIDFieldName Then Continue For
            Dim valueJson = FirestoreValueToJson(Kvp.Value)
            If valueJson IsNot Nothing Then
                Entries.Add($"""{Kvp.Key}"": {valueJson}")
            End If
        Next
        Builder.Append(String.Join(",", Entries))
        Builder.Append("}}")
        Return Builder.ToString()
    End Function
    ''' <summary>
    ''' Converts a single value into its corresponding Firestore JSON value representation.
    ''' </summary>
    ''' <param name="Value">The value to convert.</param>
    ''' <returns>A JSON string representing the value wrapped in its Firestore type.</returns>
    Private Shared Function FirestoreValueToJson(Value As Object) As String
        If Value Is Nothing Then
            Return "{""nullValue"": null}"
        End If
        If TypeOf Value Is String Then
            Return "{""stringValue"": """ & EscapeJson(Value.ToString()) & """}"
        End If
        If TypeOf Value Is Boolean Then
            Return "{""booleanValue"": " & Value.ToString().ToLower() & "}"
        End If
        If TypeOf Value Is Integer OrElse TypeOf Value Is Long Then
            Return "{""integerValue"": """ & Value.ToString() & """}"
        End If
        If TypeOf Value Is Double OrElse TypeOf Value Is Decimal OrElse TypeOf Value Is Single Then
            Dim doubleVal As Double = Convert.ToDouble(Value)
            Return "{""doubleValue"": " & doubleVal.ToString(Globalization.CultureInfo.InvariantCulture) & "}"
        End If
        If TypeOf Value Is DateTime Then
            Dim Dt As DateTime = DirectCast(Value, DateTime)
            Dim Iso = Dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            Return "{""timestampValue"": """ & Iso & """}"
        End If
        If TypeOf Value Is Dictionary(Of String, Object) Then
            Dim Dict = DirectCast(Value, Dictionary(Of String, Object))
            Dim Entries As New List(Of String)
            For Each Kvp In Dict
                Dim InnerJson = FirestoreValueToJson(Kvp.Value)
                Entries.Add($"""{Kvp.Key}"": {InnerJson}")
            Next
            Return "{""mapValue"": {""fields"": {" & String.Join(",", Entries) & "}}}"
        End If
        If TypeOf Value Is IEnumerable AndAlso TypeOf Value IsNot String Then
            Dim Values As New List(Of String)
            For Each Item In DirectCast(Value, IEnumerable)
                Values.Add(FirestoreValueToJson(Item))
            Next
            Return "{""arrayValue"": {""values"": [" & String.Join(",", Values) & "]}}"
        End If
        Return "{""stringValue"": """ & EscapeJson(Value.ToString()) & """}"
    End Function
    ''' <summary>
    ''' Escapes special characters in a string so it can be safely embedded in JSON.
    ''' </summary>
    ''' <param name="value">The string to escape.</param>
    ''' <returns>The escaped string.</returns>
    Private Shared Function EscapeJson(value As String) As String
        Return value.Replace("\", "\\").Replace("""", "\""")
    End Function
    ''' <summary>
    ''' Converts a raw Firestore document JSON response into a dictionary of field names and values.
    ''' </summary>
    ''' <param name="JsonRaw">The raw JSON string returned by Firestore.</param>
    ''' <returns>A dictionary representing the document's fields and their parsed values.</returns>
    ''' <exception cref="Exception">Thrown when the JSON cannot be converted to a dictionary.</exception>
    Private Shared Function FirestoreJsonToMap(JsonRaw As String) As Dictionary(Of String, Object)
        Dim Resultado As New Dictionary(Of String, Object)
        Try
            Using Doc As JsonDocument = JsonDocument.Parse(JsonRaw)
                Dim Root As JsonElement = Doc.RootElement
                If Root.ValueKind <> JsonValueKind.Object Then
                    Return Resultado
                End If

                ' Extrai o ID do Documento
                Dim NameProperty As JsonElement
                If Root.TryGetProperty("name", NameProperty) Then
                    Dim FullPath As String = NameProperty.GetString()
                    Dim DocumentID As String = FullPath.Split("/"c).Last()
                    Resultado.Add(DocumentIDFieldName, DocumentID)
                End If

                ' Processa os campos
                Dim FieldsProperty As JsonElement
                If Root.TryGetProperty("fields", FieldsProperty) Then
                    For Each Field As JsonProperty In FieldsProperty.EnumerateObject()
                        If Field.Name = DocumentIDFieldName Then Continue For
                        Resultado(Field.Name) = ParseFirestoreJsonElement(Field.Value)
                    Next Field
                End If
            End Using
        Catch ex As Exception
            Throw New Exception($"Error converting Firestore JSON to Dictionary.{Environment.NewLine}JSON: {JsonRaw}", ex)
        End Try
        Return Resultado
    End Function

    ''' <summary>
    ''' Parses a Firestore-typed JsonElement into its corresponding native .NET value.
    ''' </summary>
    Private Shared Function ParseFirestoreJsonElement(Element As JsonElement) As Object
        If Element.ValueKind <> JsonValueKind.Object Then Return Nothing

        For Each Prop In Element.EnumerateObject()
            Select Case Prop.Name
                Case "integerValue"
                    ' No Firestore REST API, inteiros são retornados como strings
                    If Prop.Value.ValueKind = JsonValueKind.String Then
                        Return Convert.ToInt64(Prop.Value.GetString())
                    End If
                    Return Prop.Value.GetInt64()

                Case "doubleValue"
                    If Prop.Value.ValueKind = JsonValueKind.String Then
                        Return Convert.ToDouble(Prop.Value.GetString(), Globalization.CultureInfo.InvariantCulture)
                    End If
                    Return Prop.Value.GetDouble()

                Case "booleanValue"
                    Return Prop.Value.GetBoolean()

                Case "stringValue"
                    Return Prop.Value.GetString()

                Case "timestampValue"
                    Return DateTime.Parse(Prop.Value.GetString(), Nothing, Globalization.DateTimeStyles.RoundtripKind)

                Case "nullValue"
                    Return Nothing

                Case "arrayValue"
                    Dim Result As New List(Of Object)
                    Dim ValuesElement As JsonElement
                    If Prop.Value.TryGetProperty("values", ValuesElement) AndAlso ValuesElement.ValueKind = JsonValueKind.Array Then
                        For Each Item In ValuesElement.EnumerateArray()
                            Result.Add(ParseFirestoreJsonElement(Item))
                        Next
                    End If
                    Return Result

                Case "mapValue"
                    Dim Result As New Dictionary(Of String, Object)
                    Dim FieldsElement As JsonElement
                    If Prop.Value.TryGetProperty("fields", FieldsElement) AndAlso FieldsElement.ValueKind = JsonValueKind.Object Then
                        For Each Field In FieldsElement.EnumerateObject()
                            Result(Field.Name) = ParseFirestoreJsonElement(Field.Value)
                        Next
                    End If
                    Return Result
            End Select
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Parses a Firestore-typed value object into its corresponding native .NET value.
    ''' </summary>
    ''' <param name="ValueObj">A dictionary representing the Firestore value type and its raw value.</param>
    ''' <returns>The parsed native value, which may be a primitive, list, dictionary, or <c>Nothing</c>.</returns>
    'Private Shared Function ParseFirestoreValue(ValueObj As Dictionary(Of String, Object)) As Object
    '    For Each Kvp In ValueObj
    '        Select Case Kvp.Key
    '            Case "integerValue"
    '                Return Convert.ToInt64(Kvp.Value)
    '            Case "doubleValue"
    '                Return Convert.ToDouble(Kvp.Value, Globalization.CultureInfo.InvariantCulture)
    '            Case "booleanValue"
    '                Return Convert.ToBoolean(Kvp.Value)
    '            Case "stringValue"
    '                Return Convert.ToString(Kvp.Value)
    '            Case "timestampValue"
    '                Return Convert.ToDateTime(Kvp.Value)
    '            Case "nullValue"
    '                Return Nothing
    '            Case "arrayValue"
    '                Dim Result As New List(Of Object)
    '                Dim ArrayObj = TryCast(Kvp.Value, Dictionary(Of String, Object))
    '                Dim Value As Object = Nothing
    '                If ArrayObj IsNot Nothing AndAlso ArrayObj.TryGetValue("values", Value) Then
    '                    Dim values = TryCast(Value, IEnumerable)
    '                    If values IsNot Nothing Then
    '                        For Each item In values
    '                            Dim itemDict = DirectCast(item, Dictionary(Of String, Object))
    '                            Result.Add(ParseFirestoreValue(itemDict))
    '                        Next item
    '                    End If
    '                End If
    '                Return Result
    '            Case "mapValue"
    '                Dim Result As New Dictionary(Of String, Object)
    '                Dim MapObj = DirectCast(Kvp.Value, Dictionary(Of String, Object))
    '                Dim value As Object = Nothing
    '                If MapObj.TryGetValue("fields", value) Then
    '                    Dim Fields = DirectCast(value, Dictionary(Of String, Object))
    '                    For Each Field In Fields
    '                        Dim FieldValue = DirectCast(Field.Value, Dictionary(Of String, Object))
    '                        Result(Field.Key) = ParseFirestoreValue(FieldValue)
    '                    Next Field
    '                End If
    '                Return Result
    '        End Select
    '    Next Kvp
    '    Return Nothing
    'End Function
    ''' <summary>
    ''' Converts a <see cref="FirestoreOperator"/> enumeration value into its corresponding
    ''' Firestore REST API operator string.
    ''' </summary>
    ''' <param name="Operator">The operator to convert.</param>
    ''' <returns>The string representation of the operator as expected by the Firestore API.</returns>
    Private Shared Function GetOperatorString([Operator] As FirestoreOperator) As String
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