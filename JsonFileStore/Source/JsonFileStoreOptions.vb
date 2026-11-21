Imports System
Imports System.Text.Json
Imports System.Text.Json.Serialization

''' <summary>
''' Configures JsonFileStore without exposing System.Text.Json to consumer projects.
''' </summary>
''' <remarks>
''' Options are copied into private serializer settings when the store is constructed.
''' Changes to this object after construction do not affect existing stores.
''' </remarks>
Public NotInheritable Class JsonFileStoreOptions

    ''' <summary>Writes human-readable, indented JSON.</summary>
    Public Property WriteIndented As Boolean

    ''' <summary>Matches JSON property names without regard to letter case during deserialization.</summary>
    Public Property PropertyNameCaseInsensitive As Boolean

    ''' <summary>Writes property names using camelCase instead of their declared names.</summary>
    Public Property UseCamelCasePropertyNames As Boolean

    ''' <summary>Omits properties whose values are Nothing when serializing.</summary>
    Public Property IgnoreNullValues As Boolean

    ''' <summary>Allows trailing commas when reading JSON.</summary>
    Public Property AllowTrailingCommas As Boolean

    ''' <summary>Allows // and /* */ comments when reading JSON.</summary>
    Public Property AllowJsonComments As Boolean

    ''' <summary>Serializes enumeration values as strings rather than numbers.</summary>
    Public Property SerializeEnumsAsStrings As Boolean

    ''' <summary>Ignores read-only properties when serializing objects.</summary>
    Public Property IgnoreReadOnlyProperties As Boolean

    ''' <summary>Maximum object depth. Zero uses the serializer's default.</summary>
    Public Property MaxDepth As Integer

    ''' <summary>Constructs the private serializer settings. Never expose this method publicly.</summary>
    Friend Function CreateSerializerOptions() As JsonSerializerOptions
        If MaxDepth < 0 Then Throw New ArgumentOutOfRangeException(NameOf(MaxDepth), "MaxDepth cannot be negative.")

        Dim result As New JsonSerializerOptions With {
            .WriteIndented = WriteIndented,
            .PropertyNameCaseInsensitive = PropertyNameCaseInsensitive,
            .AllowTrailingCommas = AllowTrailingCommas,
            .ReadCommentHandling = If(AllowJsonComments, JsonCommentHandling.Skip, JsonCommentHandling.Disallow),
            .IgnoreReadOnlyProperties = IgnoreReadOnlyProperties,
            .MaxDepth = MaxDepth
        }

        If UseCamelCasePropertyNames Then result.PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        If IgnoreNullValues Then result.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        If SerializeEnumsAsStrings Then result.Converters.Add(New JsonStringEnumConverter())

        Return result
    End Function
End Class
