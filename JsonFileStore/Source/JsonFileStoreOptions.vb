Imports System
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters
Imports Newtonsoft.Json.Serialization
Imports System.Reflection

''' <summary>
''' Configures JsonFileStore with library-owned JSON configuration.
''' </summary>
''' <remarks>
''' Options are copied into private serializer settings when the store is constructed.
''' Changes to this object after construction do not affect existing stores.
''' </remarks>
Public NotInheritable Class JsonFileStoreOptions
    Implements IContractResolver

    Private ReadOnly _ContractGate As New Object()
    Private _Resolver As DefaultContractResolver
    Private ReadOnly _ResolvedTypes As New HashSet(Of Type)()

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

    Friend Function CreateSerializerOptions(Optional ForReading As Boolean = False) As JsonSerializerSettings
        If MaxDepth < 0 Then Throw New ArgumentOutOfRangeException(NameOf(MaxDepth), "MaxDepth cannot be negative.")
        Dim Snapshot As JsonFileStoreOptions = CreateSnapshot()
        Snapshot._Resolver = New DefaultContractResolver()
        If UseCamelCasePropertyNames Then Snapshot._Resolver.NamingStrategy = New CamelCaseNamingStrategy(False, False)
        Dim Result As New JsonSerializerSettings With {
            .Formatting = If(WriteIndented, Formatting.Indented, Formatting.None),
            .ContractResolver = Snapshot,
            .NullValueHandling = If(IgnoreNullValues AndAlso Not ForReading, NullValueHandling.Ignore, NullValueHandling.Include),
            .ObjectCreationHandling = ObjectCreationHandling.Replace,
            .ConstructorHandling = ConstructorHandling.Default,
            .TypeNameHandling = TypeNameHandling.None,
            .MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            .DateParseHandling = DateParseHandling.None,
            .CheckAdditionalContent = True,
            .MaxDepth = If(MaxDepth = 0, 64, MaxDepth)
        }
        If SerializeEnumsAsStrings Then Result.Converters.Add(New StringEnumConverter())
        Return Result
    End Function

    Friend Function CreateSnapshot() As JsonFileStoreOptions
        Return New JsonFileStoreOptions With {
            .WriteIndented = WriteIndented,
            .PropertyNameCaseInsensitive = PropertyNameCaseInsensitive,
            .UseCamelCasePropertyNames = UseCamelCasePropertyNames,
            .IgnoreNullValues = IgnoreNullValues,
            .AllowTrailingCommas = AllowTrailingCommas,
            .AllowJsonComments = AllowJsonComments,
            .SerializeEnumsAsStrings = SerializeEnumsAsStrings,
            .IgnoreReadOnlyProperties = IgnoreReadOnlyProperties,
            .MaxDepth = MaxDepth
        }
    End Function

    Private Function ResolveContract(ObjectType As Type) As JsonContract Implements IContractResolver.ResolveContract
        SyncLock _ContractGate
            If _Resolver Is Nothing Then
                _Resolver = New DefaultContractResolver()
                If UseCamelCasePropertyNames Then _Resolver.NamingStrategy = New CamelCaseNamingStrategy(False, False)
            End If
            Dim Contract As JsonContract = _Resolver.ResolveContract(ObjectType)
            If Not _ResolvedTypes.Add(ObjectType) Then Return Contract
            Dim ObjectContract As JsonObjectContract = TryCast(Contract, JsonObjectContract)
            If ObjectContract IsNot Nothing Then
                For Each PropertyContract As JsonProperty In ObjectContract.Properties
                    If PropertyContract.DeclaringType Is Nothing OrElse PropertyContract.UnderlyingName Is Nothing Then Continue For
                    Dim Members As MemberInfo() = PropertyContract.DeclaringType.GetMember(PropertyContract.UnderlyingName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
                    Dim Field As FieldInfo = Members.OfType(Of FieldInfo)().FirstOrDefault()
                    If Field IsNot Nothing AndAlso Not Field.IsDefined(GetType(JsonPropertyAttribute), True) Then
                        PropertyContract.Ignored = True
                        Continue For
                    End If
                    Dim Member As PropertyInfo = Members.OfType(Of PropertyInfo)().FirstOrDefault()
                    If Member IsNot Nothing AndAlso Not Member.IsDefined(GetType(JsonPropertyAttribute), True) Then
                        PropertyContract.Readable = Member.GetGetMethod(False) IsNot Nothing
                        PropertyContract.Writable = Member.GetSetMethod(False) IsNot Nothing
                    End If
                    If IgnoreReadOnlyProperties AndAlso Not PropertyContract.Writable AndAlso
                        (PropertyContract.PropertyType Is GetType(String) OrElse Not GetType(System.Collections.IEnumerable).IsAssignableFrom(PropertyContract.PropertyType)) Then
                        PropertyContract.Readable = False
                    End If
                Next
            End If
            Return Contract
        End SyncLock
    End Function
End Class
