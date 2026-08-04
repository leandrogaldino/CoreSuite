Imports System.Globalization
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Friend NotInheritable Class FirestoreValueConverter
    Private Sub New()
    End Sub
    Friend Shared Function SerializeDocument(Fields As IDictionary(Of String, Object)) As JsonObject
        If Fields Is Nothing Then Throw New ArgumentNullException(NameOf(Fields))
        Dim FirestoreFields As New JsonObject()
        For Each Field As KeyValuePair(Of String, Object) In Fields
            ValidateFieldName(Field.Key)
            FirestoreFields(Field.Key) = SerializeValue(Field.Value, False)
        Next Field
        Dim Document As New JsonObject()
        Document("fields") = FirestoreFields
        Return Document
    End Function
    Friend Shared Function SerializeValue(Value As Object) As JsonObject
        Return SerializeValue(Value, False)
    End Function
    Friend Shared Function DeserializeDocument(DocumentElement As JsonElement) As FirestoreDocument
        If DocumentElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("The Firestore document response must be a JSON object.")
        Dim ResourceName As String = GetRequiredString(DocumentElement, "name")
        Dim Marker As String = "/documents/"
        Dim MarkerIndex As Integer = ResourceName.IndexOf(Marker, StringComparison.Ordinal)
        If MarkerIndex < 0 OrElse MarkerIndex + Marker.Length >= ResourceName.Length Then Throw New JsonException("The Firestore document response contains an invalid resource name.")
        Dim DocumentPath As String = ResourceName.Substring(MarkerIndex + Marker.Length)
        Dim DocumentId As String = DocumentPath.Split("/"c).Last()
        Dim Fields As New Dictionary(Of String, Object)(StringComparer.Ordinal)
        Dim FieldsElement As JsonElement
        If DocumentElement.TryGetProperty("fields", FieldsElement) Then
            If FieldsElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("The Firestore fields value must be a JSON object.")
            For Each Field As JsonProperty In FieldsElement.EnumerateObject()
                Fields(Field.Name) = DeserializeValue(Field.Value)
            Next Field
        End If
        Dim CreateTimeUtc As DateTime? = ReadTimestamp(DocumentElement, "createTime")
        Dim UpdateTimeUtc As DateTime? = ReadTimestamp(DocumentElement, "updateTime")
        Return New FirestoreDocument(DocumentId, DocumentPath, ResourceName, Fields, CreateTimeUtc, UpdateTimeUtc)
    End Function
    Friend Shared Function DeserializeValue(ValueElement As JsonElement) As Object
        If ValueElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("A Firestore value must be a JSON object.")
        For Each PropertyValue As JsonProperty In ValueElement.EnumerateObject()
            Select Case PropertyValue.Name
                Case "nullValue"
                    Return Nothing
                Case "booleanValue"
                    Return PropertyValue.Value.GetBoolean()
                Case "integerValue"
                    Return Int64.Parse(ReadStringOrRawNumber(PropertyValue.Value), NumberStyles.Integer, CultureInfo.InvariantCulture)
                Case "doubleValue"
                    Return ReadDouble(PropertyValue.Value)
                Case "timestampValue"
                    Return DateTime.Parse(PropertyValue.Value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToUniversalTime()
                Case "stringValue"
                    Return PropertyValue.Value.GetString()
                Case "bytesValue"
                    Return Convert.FromBase64String(PropertyValue.Value.GetString())
                Case "referenceValue"
                    Return New FirestoreDocumentReference(PropertyValue.Value.GetString())
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
    Private Shared Function SerializeValue(Value As Object, InsideArray As Boolean) As JsonObject
        Dim Result As New JsonObject()
        If Value Is Nothing Then
            Result("nullValue") = Nothing
            Return Result
        End If
        If TypeOf Value Is String Then
            Result("stringValue") = JsonValue.Create(DirectCast(Value, String))
            Return Result
        End If
        If TypeOf Value Is Boolean Then
            Result("booleanValue") = JsonValue.Create(DirectCast(Value, Boolean))
            Return Result
        End If
        If IsIntegralType(Value.GetType()) Then
            Result("integerValue") = JsonValue.Create(ConvertToInt64(Value).ToString(CultureInfo.InvariantCulture))
            Return Result
        End If
        If TypeOf Value Is Decimal OrElse TypeOf Value Is Double OrElse TypeOf Value Is Single Then
            Dim DoubleValue As Double = Convert.ToDouble(Value, CultureInfo.InvariantCulture)
            If Double.IsNaN(DoubleValue) OrElse Double.IsInfinity(DoubleValue) Then Throw New NotSupportedException("Non-finite floating-point values are not supported.")
            Result("doubleValue") = JsonValue.Create(DoubleValue)
            Return Result
        End If
        If TypeOf Value Is DateTime Then
            Dim DateValue As DateTime = DirectCast(Value, DateTime)
            Result("timestampValue") = JsonValue.Create(DateValue.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture))
            Return Result
        End If
        If TypeOf Value Is DateTimeOffset Then
            Result("timestampValue") = JsonValue.Create(DirectCast(Value, DateTimeOffset).UtcDateTime.ToString("O", CultureInfo.InvariantCulture))
            Return Result
        End If
        If TypeOf Value Is Byte() Then
            Result("bytesValue") = JsonValue.Create(Convert.ToBase64String(DirectCast(Value, Byte())))
            Return Result
        End If
        If TypeOf Value Is FirestoreDocumentReference Then
            Result("referenceValue") = JsonValue.Create(DirectCast(Value, FirestoreDocumentReference).ResourceName)
            Return Result
        End If
        If TypeOf Value Is FirestoreGeoPoint Then
            Dim Point As FirestoreGeoPoint = DirectCast(Value, FirestoreGeoPoint)
            Dim GeoPoint As New JsonObject()
            GeoPoint("latitude") = JsonValue.Create(Point.Latitude)
            GeoPoint("longitude") = JsonValue.Create(Point.Longitude)
            Result("geoPointValue") = GeoPoint
            Return Result
        End If
        If TypeOf Value Is IDictionary Then
            Dim MapFields As New JsonObject()
            For Each Entry As DictionaryEntry In DirectCast(Value, IDictionary)
                Dim FieldName As String = TryCast(Entry.Key, String)
                If FieldName Is Nothing Then Throw New NotSupportedException("Firestore map keys must be strings.")
                ValidateFieldName(FieldName)
                MapFields(FieldName) = SerializeValue(Entry.Value, False)
            Next Entry
            Dim MapValue As New JsonObject()
            MapValue("fields") = MapFields
            Result("mapValue") = MapValue
            Return Result
        End If
        If TypeOf Value Is IEnumerable Then
            If InsideArray Then Throw New NotSupportedException("Firestore arrays cannot directly contain another array.")
            Dim Values As New JsonArray()
            For Each Item As Object In DirectCast(Value, IEnumerable)
                Values.Add(SerializeValue(Item, True))
            Next Item
            Dim ArrayValue As New JsonObject()
            ArrayValue("values") = Values
            Result("arrayValue") = ArrayValue
            Return Result
        End If
        Throw New NotSupportedException($"The .NET type '{Value.GetType().FullName}' cannot be stored as a Firestore value.")
    End Function
    Private Shared Function DeserializeArray(ArrayElement As JsonElement) As List(Of Object)
        Dim Result As New List(Of Object)()
        If ArrayElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("A Firestore arrayValue must be a JSON object.")
        Dim ValuesElement As JsonElement
        If Not ArrayElement.TryGetProperty("values", ValuesElement) Then Return Result
        If ValuesElement.ValueKind <> JsonValueKind.Array Then Throw New JsonException("The values property of arrayValue must be a JSON array.")
        For Each Item As JsonElement In ValuesElement.EnumerateArray()
            Result.Add(DeserializeValue(Item))
        Next Item
        Return Result
    End Function
    Private Shared Function DeserializeMap(MapElement As JsonElement) As Dictionary(Of String, Object)
        Dim Result As New Dictionary(Of String, Object)(StringComparer.Ordinal)
        If MapElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("A Firestore mapValue must be a JSON object.")
        Dim FieldsElement As JsonElement
        If Not MapElement.TryGetProperty("fields", FieldsElement) Then Return Result
        If FieldsElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("The fields property of mapValue must be a JSON object.")
        For Each Field As JsonProperty In FieldsElement.EnumerateObject()
            Result(Field.Name) = DeserializeValue(Field.Value)
        Next Field
        Return Result
    End Function
    Private Shared Function DeserializeGeoPoint(GeoPointElement As JsonElement) As FirestoreGeoPoint
        If GeoPointElement.ValueKind <> JsonValueKind.Object Then Throw New JsonException("A Firestore geoPointValue must be a JSON object.")
        Dim LatitudeElement As JsonElement
        Dim LongitudeElement As JsonElement
        If Not GeoPointElement.TryGetProperty("latitude", LatitudeElement) OrElse Not GeoPointElement.TryGetProperty("longitude", LongitudeElement) Then Throw New JsonException("The Firestore geoPointValue is incomplete.")
        Return New FirestoreGeoPoint(LatitudeElement.GetDouble(), LongitudeElement.GetDouble())
    End Function
    Private Shared Function ReadDouble(ValueElement As JsonElement) As Double
        If ValueElement.ValueKind = JsonValueKind.Number Then Return ValueElement.GetDouble()
        If ValueElement.ValueKind <> JsonValueKind.String Then Throw New JsonException("The Firestore doubleValue is invalid.")
        Select Case ValueElement.GetString()
            Case "NaN"
                Return Double.NaN
            Case "Infinity"
                Return Double.PositiveInfinity
            Case "-Infinity"
                Return Double.NegativeInfinity
            Case Else
                Return Double.Parse(ValueElement.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture)
        End Select
    End Function
    Private Shared Function ReadStringOrRawNumber(ValueElement As JsonElement) As String
        If ValueElement.ValueKind = JsonValueKind.String Then Return ValueElement.GetString()
        If ValueElement.ValueKind = JsonValueKind.Number Then Return ValueElement.GetRawText()
        Throw New JsonException("The Firestore integerValue is invalid.")
    End Function
    Private Shared Function ReadTimestamp(DocumentElement As JsonElement, PropertyName As String) As DateTime?
        Dim TimestampElement As JsonElement
        If Not DocumentElement.TryGetProperty(PropertyName, TimestampElement) Then Return Nothing
        If TimestampElement.ValueKind <> JsonValueKind.String Then Throw New JsonException($"The Firestore {PropertyName} value is invalid.")
        Return DateTime.Parse(TimestampElement.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToUniversalTime()
    End Function
    Private Shared Function GetRequiredString(Element As JsonElement, PropertyName As String) As String
        Dim PropertyElement As JsonElement
        If Not Element.TryGetProperty(PropertyName, PropertyElement) OrElse PropertyElement.ValueKind <> JsonValueKind.String OrElse String.IsNullOrWhiteSpace(PropertyElement.GetString()) Then Throw New JsonException($"The required Firestore property '{PropertyName}' is missing or invalid.")
        Return PropertyElement.GetString()
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
