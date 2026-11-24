Imports System.Globalization
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Friend NotInheritable Class FirestoreValueConverter
    Private Sub New()
    End Sub
    Friend Shared Function SerializeDocument(Fields As IDictionary(Of String, Object)) As JObject
        If Fields Is Nothing Then Throw New ArgumentNullException(NameOf(Fields))
        Dim FirestoreFields As New JObject()
        For Each Field As KeyValuePair(Of String, Object) In Fields
            ValidateFieldName(Field.Key)
            FirestoreFields(Field.Key) = SerializeValue(Field.Value, False)
        Next Field
        Dim Document As New JObject()
        Document("fields") = FirestoreFields
        Return Document
    End Function
    Friend Shared Function SerializeValue(Value As Object) As JObject
        Return SerializeValue(Value, False)
    End Function
    Friend Shared Function DeserializeDocument(DocumentElement As JToken) As FirestoreDocument
        If DocumentElement.Type <> JTokenType.Object Then Throw New JsonException("The Firestore document response must be a JSON object.")
        Dim ResourceName As String = GetRequiredString(DocumentElement, "name")
        Dim Marker As String = "/documents/"
        Dim MarkerIndex As Integer = ResourceName.IndexOf(Marker, StringComparison.Ordinal)
        If MarkerIndex < 0 OrElse MarkerIndex + Marker.Length >= ResourceName.Length Then Throw New JsonException("The Firestore document response contains an invalid resource name.")
        Dim DocumentPath As String = ResourceName.Substring(MarkerIndex + Marker.Length)
        Dim DocumentId As String = DocumentPath.Split("/"c).Last()
        Dim Fields As New Dictionary(Of String, Object)(StringComparer.Ordinal)
        Dim FieldsElement As JToken = Nothing
        If FirebaseClient.TryGetJsonProperty(DocumentElement, "fields", FieldsElement) Then
            If FieldsElement.Type <> JTokenType.Object Then Throw New JsonException("The Firestore fields value must be a JSON object.")
            For Each Field As JProperty In DirectCast(FieldsElement, JObject).Properties()
                Fields(Field.Name) = DeserializeValue(Field.Value)
            Next Field
        End If
        Dim CreateTimeUtc As DateTime? = ReadTimestamp(DocumentElement, "createTime")
        Dim UpdateTimeUtc As DateTime? = ReadTimestamp(DocumentElement, "updateTime")
        Return New FirestoreDocument(DocumentId, DocumentPath, ResourceName, Fields, CreateTimeUtc, UpdateTimeUtc)
    End Function
    Friend Shared Function DeserializeValue(ValueElement As JToken) As Object
        If ValueElement.Type <> JTokenType.Object Then Throw New JsonException("A Firestore value must be a JSON object.")
        For Each PropertyValue As JProperty In DirectCast(ValueElement, JObject).Properties()
            Select Case PropertyValue.Name
                Case "nullValue"
                    Return Nothing
                Case "booleanValue"
                    Return FirebaseClient.ReadJsonBoolean(PropertyValue.Value)
                Case "integerValue"
                    Return Int64.Parse(ReadStringOrRawNumber(PropertyValue.Value), NumberStyles.Integer, CultureInfo.InvariantCulture)
                Case "doubleValue"
                    Return ReadDouble(PropertyValue.Value)
                Case "timestampValue"
                    Return DateTime.Parse(FirebaseClient.ReadJsonString(PropertyValue.Value), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToUniversalTime()
                Case "stringValue"
                    Return FirebaseClient.ReadJsonString(PropertyValue.Value)
                Case "bytesValue"
                    Return Convert.FromBase64String(FirebaseClient.ReadJsonString(PropertyValue.Value))
                Case "referenceValue"
                    Return New FirestoreDocumentReference(FirebaseClient.ReadJsonString(PropertyValue.Value))
                Case "geoPointValue"
                    Return DeserializeGeoPoint(PropertyValue.Value)
                Case "arrayValue"
                    Return DeserializeArray(PropertyValue.Value)
                Case "mapValue"
                    Return DeserializeMap(PropertyValue.Value)
            End Select
        Next PropertyValue
        Throw New JsonException("The Firestore value contains an unsupported or missing value type.")
    End Function
    Private Shared Function SerializeValue(Value As Object, InsideArray As Boolean) As JObject
        Dim Result As New JObject()
        If Value Is Nothing Then
            Result("nullValue") = JValue.CreateNull()
            Return Result
        End If
        If TypeOf Value Is String Then
            Result("stringValue") = New JValue(DirectCast(Value, String))
            Return Result
        End If
        If TypeOf Value Is Boolean Then
            Result("booleanValue") = New JValue(DirectCast(Value, Boolean))
            Return Result
        End If
        If IsIntegralType(Value.GetType()) Then
            Result("integerValue") = New JValue(ConvertToInt64(Value).ToString(CultureInfo.InvariantCulture))
            Return Result
        End If
        If TypeOf Value Is Decimal OrElse TypeOf Value Is Double OrElse TypeOf Value Is Single Then
            Dim DoubleValue As Double = Convert.ToDouble(Value, CultureInfo.InvariantCulture)
            If Double.IsNaN(DoubleValue) OrElse Double.IsInfinity(DoubleValue) Then Throw New NotSupportedException("Non-finite floating-point values are not supported.")
            Result("doubleValue") = New JValue(DoubleValue)
            Return Result
        End If
        If TypeOf Value Is DateTime Then
            Dim DateValue As DateTime = DirectCast(Value, DateTime)
            Result("timestampValue") = New JValue(DateValue.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture))
            Return Result
        End If
        If TypeOf Value Is DateTimeOffset Then
            Result("timestampValue") = New JValue(DirectCast(Value, DateTimeOffset).UtcDateTime.ToString("O", CultureInfo.InvariantCulture))
            Return Result
        End If
        If TypeOf Value Is Byte() Then
            Result("bytesValue") = New JValue(Convert.ToBase64String(DirectCast(Value, Byte())))
            Return Result
        End If
        If TypeOf Value Is FirestoreDocumentReference Then
            Result("referenceValue") = New JValue(DirectCast(Value, FirestoreDocumentReference).ResourceName)
            Return Result
        End If
        If TypeOf Value Is FirestoreGeoPoint Then
            Dim Point As FirestoreGeoPoint = DirectCast(Value, FirestoreGeoPoint)
            Dim GeoPoint As New JObject()
            GeoPoint("latitude") = New JValue(Point.Latitude)
            GeoPoint("longitude") = New JValue(Point.Longitude)
            Result("geoPointValue") = GeoPoint
            Return Result
        End If
        If TypeOf Value Is IDictionary Then
            Dim MapFields As New JObject()
            For Each Entry As DictionaryEntry In DirectCast(Value, IDictionary)
                Dim FieldName As String = TryCast(Entry.Key, String)
                If FieldName Is Nothing Then Throw New NotSupportedException("Firestore map keys must be strings.")
                ValidateFieldName(FieldName)
                MapFields(FieldName) = SerializeValue(Entry.Value, False)
            Next Entry
            Dim MapValue As New JObject()
            MapValue("fields") = MapFields
            Result("mapValue") = MapValue
            Return Result
        End If
        If TypeOf Value Is IEnumerable Then
            If InsideArray Then Throw New NotSupportedException("Firestore arrays cannot directly contain another array.")
            Dim Values As New JArray()
            For Each Item As Object In DirectCast(Value, IEnumerable)
                Values.Add(SerializeValue(Item, True))
            Next Item
            Dim ArrayValue As New JObject()
            ArrayValue("values") = Values
            Result("arrayValue") = ArrayValue
            Return Result
        End If
        Throw New NotSupportedException($"The .NET type '{Value.GetType().FullName}' cannot be stored as a Firestore value.")
    End Function
    Private Shared Function DeserializeArray(ArrayElement As JToken) As List(Of Object)
        Dim Result As New List(Of Object)()
        If ArrayElement.Type <> JTokenType.Object Then Throw New JsonException("A Firestore arrayValue must be a JSON object.")
        Dim ValuesElement As JToken = Nothing
        If Not FirebaseClient.TryGetJsonProperty(ArrayElement, "values", ValuesElement) Then Return Result
        If ValuesElement.Type <> JTokenType.Array Then Throw New JsonException("The values property of arrayValue must be a JSON array.")
        For Each Item As JToken In ValuesElement.Children()
            Result.Add(DeserializeValue(Item))
        Next Item
        Return Result
    End Function
    Private Shared Function DeserializeMap(MapElement As JToken) As Dictionary(Of String, Object)
        Dim Result As New Dictionary(Of String, Object)(StringComparer.Ordinal)
        If MapElement.Type <> JTokenType.Object Then Throw New JsonException("A Firestore mapValue must be a JSON object.")
        Dim FieldsElement As JToken = Nothing
        If Not FirebaseClient.TryGetJsonProperty(MapElement, "fields", FieldsElement) Then Return Result
        If FieldsElement.Type <> JTokenType.Object Then Throw New JsonException("The fields property of mapValue must be a JSON object.")
        For Each Field As JProperty In DirectCast(FieldsElement, JObject).Properties()
            Result(Field.Name) = DeserializeValue(Field.Value)
        Next Field
        Return Result
    End Function
    Private Shared Function DeserializeGeoPoint(GeoPointElement As JToken) As FirestoreGeoPoint
        If GeoPointElement.Type <> JTokenType.Object Then Throw New JsonException("A Firestore geoPointValue must be a JSON object.")
        Dim LatitudeElement As JToken = Nothing
        Dim LongitudeElement As JToken = Nothing
        If Not FirebaseClient.TryGetJsonProperty(GeoPointElement, "latitude", LatitudeElement) OrElse Not FirebaseClient.TryGetJsonProperty(GeoPointElement, "longitude", LongitudeElement) Then Throw New JsonException("The Firestore geoPointValue is incomplete.")
        Return New FirestoreGeoPoint(FirebaseClient.ReadJsonDouble(LatitudeElement), FirebaseClient.ReadJsonDouble(LongitudeElement))
    End Function
    Private Shared Function ReadDouble(ValueElement As JToken) As Double
        If FirebaseClient.IsJsonNumber(ValueElement) Then Return FirebaseClient.ReadJsonDouble(ValueElement)
        If ValueElement.Type <> JTokenType.String Then Throw New JsonException("The Firestore doubleValue is invalid.")
        Select Case FirebaseClient.ReadJsonString(ValueElement)
            Case "NaN"
                Return Double.NaN
            Case "Infinity"
                Return Double.PositiveInfinity
            Case "-Infinity"
                Return Double.NegativeInfinity
            Case Else
                Return Double.Parse(FirebaseClient.ReadJsonString(ValueElement), NumberStyles.Float, CultureInfo.InvariantCulture)
        End Select
    End Function
    Private Shared Function ReadStringOrRawNumber(ValueElement As JToken) As String
        If ValueElement.Type = JTokenType.String Then Return FirebaseClient.ReadJsonString(ValueElement)
        If FirebaseClient.IsJsonNumber(ValueElement) Then Return ValueElement.ToString(Formatting.None)
        Throw New JsonException("The Firestore integerValue is invalid.")
    End Function
    Private Shared Function ReadTimestamp(DocumentElement As JToken, PropertyName As String) As DateTime?
        Dim TimestampElement As JToken = Nothing
        If Not FirebaseClient.TryGetJsonProperty(DocumentElement, PropertyName, TimestampElement) Then Return Nothing
        If TimestampElement.Type <> JTokenType.String Then Throw New JsonException($"The Firestore {PropertyName} value is invalid.")
        Return DateTime.Parse(FirebaseClient.ReadJsonString(TimestampElement), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToUniversalTime()
    End Function
    Private Shared Function GetRequiredString(Element As JToken, PropertyName As String) As String
        Dim PropertyElement As JToken = Nothing
        If Not FirebaseClient.TryGetJsonProperty(Element, PropertyName, PropertyElement) OrElse PropertyElement.Type <> JTokenType.String OrElse String.IsNullOrWhiteSpace(FirebaseClient.ReadJsonString(PropertyElement)) Then Throw New JsonException($"The required Firestore property '{PropertyName}' is missing or invalid.")
        Return FirebaseClient.ReadJsonString(PropertyElement)
    End Function
    Private Shared Function IsIntegralType(ValueType As Type) As Boolean
        Select Case System.Type.GetTypeCode(ValueType)
            Case TypeCode.SByte, TypeCode.Byte, TypeCode.Int16, TypeCode.UInt16, TypeCode.Int32, TypeCode.UInt32, TypeCode.Int64, TypeCode.UInt64
                Return True
            Case Else
                Return False
        End Select
    End Function
    Private Shared Function ConvertToInt64(Value As Object) As Long
        If TypeOf Value Is ULong AndAlso DirectCast(Value, ULong) > Long.MaxValue Then Throw New OverflowException("The unsigned integer is larger than the maximum Firestore integer value.")
        Return Convert.ToInt64(Value, CultureInfo.InvariantCulture)
    End Function
    Private Shared Sub ValidateFieldName(Value As String)
        If String.IsNullOrEmpty(Value) Then Throw New ArgumentException("Firestore field names cannot be empty.", NameOf(Value))
        If Encoding.UTF8.GetByteCount(Value) > 1500 Then Throw New ArgumentException("Firestore field names cannot exceed 1,500 UTF-8 bytes.", NameOf(Value))
        If Value.StartsWith("__", StringComparison.Ordinal) AndAlso Value.EndsWith("__", StringComparison.Ordinal) Then Throw New ArgumentException("Firestore field names matching '__.*__' are reserved.", NameOf(Value))
    End Sub
End Class
