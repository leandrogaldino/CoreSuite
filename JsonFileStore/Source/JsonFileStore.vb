Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports System.Text
Imports System.Threading
''' <summary>
''' Provides strongly typed JSON persistence with atomic writes, backup creation, and automatic recovery.
''' </summary>
''' <typeparam name="T">The type of value stored in the JSON file.</typeparam>
''' <remarks>
''' Operations performed through the same instance are serialized. A successful save replaces the primary file only after
''' serialization has completed, while the previous valid primary file is retained as the backup.
''' </remarks>
Public NotInheritable Class JsonFileStore(Of T)
    Private Const FileBufferSize As Integer = 81920
    Private ReadOnly _FilePath As String
    Private ReadOnly _BackupPath As String
    Private ReadOnly _OperationGate As New SemaphoreSlim(1, 1)
    Private ReadOnly _SerializerOptions As JsonSerializerSettings
    Private ReadOnly _ReaderOptions As JsonSerializerSettings
    Private ReadOnly _Options As JsonFileStoreOptions
    Private _AutoRecoverFromBackup As Boolean = True
    ''' <summary>
    ''' Initializes a new instance of the <see cref="JsonFileStore(Of T)"/> class.
    ''' </summary>
    ''' <param name="FilePath">The path of the primary JSON file.</param>
    ''' <param name="Options">Optional, library-owned JSON settings. They are snapshotted on construction.</param>
    ''' <param name="BackupPath">An optional backup path. The default is the primary path followed by <c>.bak</c>.</param>
    ''' <exception cref="ArgumentNullException"><paramref name="FilePath"/> is <see langword="Nothing"/>.</exception>
    ''' <exception cref="ArgumentException">A path is empty, contains only whitespace, or both paths resolve to the same file.</exception>
    Public Sub New(FilePath As String, Optional Options As JsonFileStoreOptions = Nothing, Optional BackupPath As String = Nothing)
        If FilePath Is Nothing Then Throw New ArgumentNullException("FilePath")
        If String.IsNullOrWhiteSpace(FilePath) Then Throw New ArgumentException("The JSON file path cannot be empty or whitespace.", NameOf(FilePath))
        _FilePath = Path.GetFullPath(FilePath)
        If BackupPath Is Nothing Then
            _BackupPath = _FilePath & ".bak"
        Else
            If String.IsNullOrWhiteSpace(BackupPath) Then Throw New ArgumentException("The backup file path cannot be empty or whitespace.", NameOf(BackupPath))
            _BackupPath = Path.GetFullPath(BackupPath)
        End If
        Dim pathComparison As StringComparison = If((Environment.OSVersion.Platform = PlatformID.Win32NT), StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
        If String.Equals(_FilePath, _BackupPath, pathComparison) Then Throw New ArgumentException("The primary and backup paths must be different.", NameOf(BackupPath))
        _Options = If(Options, New JsonFileStoreOptions()).CreateSnapshot()
        _SerializerOptions = _Options.CreateSerializerOptions()
        _ReaderOptions = _Options.CreateSerializerOptions(True)
    End Sub
    ''' <summary>
    ''' Gets the absolute path of the primary JSON file.
    ''' </summary>
    ''' <value>The absolute primary-file path.</value>
    Public ReadOnly Property FilePath As String
        Get
            Return _FilePath
        End Get
    End Property
    ''' <summary>
    ''' Gets the absolute path of the backup file.
    ''' </summary>
    ''' <value>The absolute backup-file path.</value>
    Public ReadOnly Property BackupPath As String
        Get
            Return _BackupPath
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets whether a failed primary load should automatically use the backup and restore the primary file.
    ''' </summary>
    ''' <value><see langword="True"/> to recover automatically; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    Public Property AutoRecoverFromBackup As Boolean
        Get
            Return _AutoRecoverFromBackup
        End Get
        Set(value As Boolean)
            _AutoRecoverFromBackup = value
        End Set
    End Property
    ''' <summary>
    ''' Gets whether the primary JSON file currently exists.
    ''' </summary>
    ''' <value><see langword="True"/> when the primary file exists; otherwise, <see langword="False"/>.</value>
    Public ReadOnly Property Exists As Boolean
        Get
            Return File.Exists(_FilePath)
        End Get
    End Property
    ''' <summary>
    ''' Gets whether the backup file currently exists.
    ''' </summary>
    ''' <value><see langword="True"/> when the backup exists; otherwise, <see langword="False"/>.</value>
    Public ReadOnly Property BackupExists As Boolean
        Get
            Return File.Exists(_BackupPath)
        End Get
    End Property
    ''' <summary>
    ''' Serializes a value and atomically saves it to the primary JSON file.
    ''' </summary>
    ''' <param name="Value">The value to serialize. Reference types may be <see langword="Nothing"/>.</param>
    Public Sub Save(Value As T)
        _OperationGate.Wait()
        Try
            SaveCore(Value)
        Finally
            _OperationGate.Release()
        End Try
    End Sub
    ''' <summary>
    ''' Asynchronously serializes a value and atomically saves it to the primary JSON file.
    ''' </summary>
    ''' <param name="Value">The value to serialize. Reference types may be <see langword="Nothing"/>.</param>
    ''' <param name="CancellationToken">A token used to cancel serialization and file I/O before the atomic commit.</param>
    ''' <returns>A task that represents the save operation.</returns>
    Public Async Function SaveAsync(Value As T, Optional CancellationToken As CancellationToken = Nothing) As Task
        Await _OperationGate.WaitAsync(CancellationToken).ConfigureAwait(False)
        Try
            Await SaveCoreAsync(Value, CancellationToken).ConfigureAwait(False)
        Finally
            _OperationGate.Release()
        End Try
    End Function
    ''' <summary>
    ''' Loads the stored value, automatically recovering from the backup when enabled and necessary.
    ''' </summary>
    ''' <returns>The deserialized value.</returns>
    ''' <exception cref="IOException">The primary file could not be read and no backup was available.</exception>
    ''' <exception cref="JsonException">The primary JSON was invalid and no backup was available.</exception>
    ''' <exception cref="JsonFileRecoveryException">A backup was available, but the recovery operation failed.</exception>
    Public Function Load() As T
        _OperationGate.Wait()
        Try
            Return LoadCore()
        Finally
            _OperationGate.Release()
        End Try
    End Function
    ''' <summary>
    ''' Asynchronously loads the stored value, automatically recovering from the backup when enabled and necessary.
    ''' </summary>
    ''' <param name="CancellationToken">A token used to cancel file I/O.</param>
    ''' <returns>A task whose result is the deserialized value.</returns>
    ''' <exception cref="IOException">The primary file could not be read and no backup was available.</exception>
    ''' <exception cref="JsonException">The primary JSON was invalid and no backup was available.</exception>
    ''' <exception cref="JsonFileRecoveryException">A backup was available, but the recovery operation failed.</exception>
    Public Async Function LoadAsync(Optional CancellationToken As CancellationToken = Nothing) As Task(Of T)
        Await _OperationGate.WaitAsync(CancellationToken).ConfigureAwait(False)
        Try
            Return Await LoadCoreAsync(CancellationToken).ConfigureAwait(False)
        Finally
            _OperationGate.Release()
        End Try
    End Function
    ''' <summary>
    ''' Attempts to load the stored value without propagating expected file, access, or JSON errors.
    ''' </summary>
    ''' <param name="Value">Receives the loaded value, or the default value of <typeparamref name="T"/> when loading fails.</param>
    ''' <returns><see langword="True"/> when a value was loaded; otherwise, <see langword="False"/>.</returns>
    Public Function TryLoad(ByRef Value As T) As Boolean
        Try
            Value = Load()
            Return True
        Catch ex As Exception When IsExpectedLoadException(ex)
            Value = Nothing
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Loads the stored value or returns a caller-provided default when loading and recovery fail.
    ''' </summary>
    ''' <param name="DefaultValue">The value returned when no stored value can be loaded.</param>
    ''' <returns>The stored value, or <paramref name="DefaultValue"/>.</returns>
    Public Function LoadOrDefault(Optional DefaultValue As T = Nothing) As T
        Try
            Return Load()
        Catch ex As Exception When IsExpectedLoadException(ex)
            Return DefaultValue
        End Try
    End Function
    ''' <summary>
    ''' Asynchronously loads the stored value or returns a caller-provided default when loading and recovery fail.
    ''' </summary>
    ''' <param name="DefaultValue">The value returned when no stored value can be loaded.</param>
    ''' <param name="CancellationToken">A token used to cancel file I/O.</param>
    ''' <returns>A task whose result is the stored value, or <paramref name="DefaultValue"/>.</returns>
    Public Async Function LoadOrDefaultAsync(Optional DefaultValue As T = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of T)
        Try
            Return Await LoadAsync(CancellationToken).ConfigureAwait(False)
        Catch ex As Exception When IsExpectedLoadException(ex)
            Return DefaultValue
        End Try
    End Function
    Private Sub SaveCore(Value As T)
        EnsureParentDirectory(_FilePath)
        EnsureParentDirectory(_BackupPath)
        Dim TemporaryPath As String = CreateTemporaryPath(_FilePath)
        Try
            SerializeToFile(TemporaryPath, Value)
            If File.Exists(_FilePath) AndAlso IsValidStoredValue(_FilePath) Then
                CopyFileAtomically(_FilePath, _BackupPath)
            End If
            ' Finish backup preparation before committing the new primary file.
            If Not File.Exists(_BackupPath) Then CopyFileAtomically(TemporaryPath, _BackupPath)
            PromoteTemporaryFile(TemporaryPath, _FilePath)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Sub
    Private Async Function SaveCoreAsync(Value As T, CancellationToken As CancellationToken) As Task
        EnsureParentDirectory(_FilePath)
        EnsureParentDirectory(_BackupPath)
        Dim TemporaryPath As String = CreateTemporaryPath(_FilePath)
        Try
            Await SerializeToFileAsync(TemporaryPath, Value, CancellationToken).ConfigureAwait(False)
            If File.Exists(_FilePath) AndAlso Await IsValidStoredValueAsync(_FilePath, CancellationToken).ConfigureAwait(False) Then
                Await CopyFileAtomicallyAsync(_FilePath, _BackupPath, CancellationToken).ConfigureAwait(False)
            End If
            If Not File.Exists(_BackupPath) Then
                Await CopyFileAtomicallyAsync(TemporaryPath, _BackupPath, CancellationToken).ConfigureAwait(False)
            End If
            CancellationToken.ThrowIfCancellationRequested()
            PromoteTemporaryFile(TemporaryPath, _FilePath)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Function
    Private Function LoadCore() As T
        Dim PrimaryException As Exception
        Try
            Return DeserializeFromFile(_FilePath)
        Catch ex As Exception When IsRecoverablePrimaryFailure(ex)
            PrimaryException = ex
        End Try
        If Not _AutoRecoverFromBackup OrElse Not File.Exists(_BackupPath) Then Throw PrimaryException
        Try
            Dim RecoveredValue As T = DeserializeFromFile(_BackupPath)
            CopyFileAtomically(_BackupPath, _FilePath)
            Return RecoveredValue
        Catch RecoveryException As Exception When IsExpectedLoadException(RecoveryException)
            Throw New JsonFileRecoveryException(_FilePath, _BackupPath, PrimaryException, RecoveryException)
        End Try
    End Function
    Private Async Function LoadCoreAsync(CancellationToken As CancellationToken) As Task(Of T)
        Dim PrimaryException As Exception
        Try
            Return Await DeserializeFromFileAsync(_FilePath, CancellationToken).ConfigureAwait(False)
        Catch ex As Exception When IsRecoverablePrimaryFailure(ex)
            PrimaryException = ex
        End Try
        If Not _AutoRecoverFromBackup OrElse Not File.Exists(_BackupPath) Then Throw PrimaryException
        Try
            Dim RecoveredValue As T = Await DeserializeFromFileAsync(_BackupPath, CancellationToken).ConfigureAwait(False)
            Await CopyFileAtomicallyAsync(_BackupPath, _FilePath, CancellationToken).ConfigureAwait(False)
            Return RecoveredValue
        Catch recoveryException As Exception When IsExpectedLoadException(recoveryException)
            Throw New JsonFileRecoveryException(_FilePath, _BackupPath, PrimaryException, recoveryException)
        End Try
    End Function
    Private Sub SerializeToFile(FilePath As String, Value As T)
        Using OutputStream As New FileStream(FilePath, FileMode.CreateNew, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.WriteThrough Or FileOptions.SequentialScan)
            Using TextWriter As New StreamWriter(OutputStream, New UTF8Encoding(False, True), FileBufferSize, True)
                Using Writer As New JsonTextWriter(TextWriter) With {.CloseOutput = False}
                    JsonSerializer.Create(_SerializerOptions).Serialize(Writer, Value, GetType(T))
                    Writer.Flush()
                End Using
            End Using
            OutputStream.Flush(True)
        End Using
    End Sub
    Private Async Function SerializeToFileAsync(FilePath As String, Value As T, CancellationToken As CancellationToken) As Task
        ' Json.NET's object serializer is synchronous; keep it off the caller's thread.
        Await Task.Run(Sub() SerializeToFile(FilePath, Value), CancellationToken).ConfigureAwait(False)
        CancellationToken.ThrowIfCancellationRequested()
    End Function
    Private Function DeserializeFromFile(FilePath As String) As T
        Using InputStream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.SequentialScan)
            Using TextReader As New StreamReader(InputStream, New UTF8Encoding(False, True), False)
                Return DeserializeJson(TextReader.ReadToEnd(), CancellationToken.None)
            End Using
        End Using
    End Function
    Private Async Function DeserializeFromFileAsync(FilePath As String, CancellationToken As CancellationToken) As Task(Of T)
        Dim Content As String
        Using InputStream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.Asynchronous Or FileOptions.SequentialScan)
            Using TextReader As New StreamReader(InputStream, New UTF8Encoding(False, True), False)
                Dim Buffer(8191) As Char
                Dim Builder As New StringBuilder()
                Do
                    CancellationToken.ThrowIfCancellationRequested()
                    Dim Count As Integer = Await TextReader.ReadAsync(Buffer, 0, Buffer.Length).ConfigureAwait(False)
                    If Count = 0 Then Exit Do
                    Builder.Append(Buffer, 0, Count)
                Loop
                Content = Builder.ToString()
            End Using
        End Using
        Dim Result As T = Await Task.Run(Function() DeserializeJson(Content, CancellationToken), CancellationToken).ConfigureAwait(False)
        CancellationToken.ThrowIfCancellationRequested()
        Return Result
    End Function
    Private Function DeserializeJson(Content As String, Token As CancellationToken) As T
        If Content.Length > 0 AndAlso Content(0) = ChrW(&HFEFF) Then Content = Content.Substring(1)
        ValidateJsonSyntax(Content, Token)
        Using TextReader As New StringReader(Content)
            Using Reader As New JsonTextReader(TextReader) With {.DateParseHandling = DateParseHandling.None, .MaxDepth = _ReaderOptions.MaxDepth}
                While Reader.Read() AndAlso Reader.TokenType = JsonToken.Comment
                End While
                If Reader.TokenType = JsonToken.None OrElse Reader.TokenType = JsonToken.Comment Then Throw New JsonReaderException("The JSON value is missing.")
                Dim Value As JToken = JToken.ReadFrom(Reader)
                While Reader.Read()
                    If Reader.TokenType <> JsonToken.Comment Then Throw New JsonReaderException("Unexpected content after the JSON value.")
                End While
                If Not _Options.PropertyNameCaseInsensitive Then FilterCaseMismatches(Value, GetType(T), Token)
                Token.ThrowIfCancellationRequested()
                Return Value.ToObject(Of T)(JsonSerializer.Create(_ReaderOptions))
            End Using
        End Using
    End Function
    Private Sub FilterCaseMismatches(Value As JToken, TargetType As Type, Token As CancellationToken)
        Token.ThrowIfCancellationRequested()
        If Value Is Nothing OrElse TargetType Is Nothing OrElse TargetType Is GetType(Object) Then Return
        TargetType = If(Nullable.GetUnderlyingType(TargetType), TargetType)
        Dim Contract As JsonContract = _ReaderOptions.ContractResolver.ResolveContract(TargetType)
        If Contract.Converter IsNot Nothing Then Return
        Dim ObjectValue As JObject = TryCast(Value, JObject)
        Dim ObjectContract As JsonObjectContract = TryCast(Contract, JsonObjectContract)
        If ObjectValue IsNot Nothing AndAlso ObjectContract IsNot Nothing Then
            For Each Member As JProperty In ObjectValue.Properties().ToList()
                Dim Exact As JsonProperty = ObjectContract.Properties.GetProperty(Member.Name, StringComparison.Ordinal)
                If Exact Is Nothing Then
                    Dim Similar As JsonProperty = ObjectContract.Properties.GetProperty(Member.Name, StringComparison.OrdinalIgnoreCase)
                    If Similar IsNot Nothing Then Member.Remove()
                ElseIf Not Exact.Ignored AndAlso Exact.Converter Is Nothing Then
                    FilterCaseMismatches(Member.Value, Exact.PropertyType, Token)
                End If
            Next
        ElseIf ObjectValue IsNot Nothing AndAlso TypeOf Contract Is JsonDictionaryContract Then
            Dim ItemType As Type = DirectCast(Contract, JsonDictionaryContract).DictionaryValueType
            For Each Member As JProperty In ObjectValue.Properties()
                FilterCaseMismatches(Member.Value, ItemType, Token)
            Next
        ElseIf TypeOf Value Is JArray AndAlso TypeOf Contract Is JsonArrayContract Then
            Dim ItemType As Type = DirectCast(Contract, JsonArrayContract).CollectionItemType
            For Each Item As JToken In Value.Children()
                FilterCaseMismatches(Item, ItemType, Token)
            Next
        End If
    End Sub
    Private Sub ValidateJsonSyntax(Content As String, Token As CancellationToken)
        ' Json.NET accepts comments and trailing commas by default; enforce the store's existing options.
        Dim InString As Boolean = False
        Dim Escaped As Boolean = False
        Dim Previous As Char = ChrW(0)
        Dim Index As Integer = 0
        While Index < Content.Length
            If (Index And 4095) = 0 Then Token.ThrowIfCancellationRequested()
            Dim Character As Char = Content(Index)
            If InString Then
                If Escaped Then
                    Escaped = False
                ElseIf Character = "\"c Then
                    Escaped = True
                ElseIf Character = """"c Then
                    InString = False
                    Previous = Character
                ElseIf Character < ChrW(32) Then
                    Throw New JsonReaderException("Unescaped control character in JSON string.")
                End If
            ElseIf Character = """"c Then
                InString = True
            ElseIf Character = "'"c Then
                Throw New JsonReaderException("JSON strings must use double quotes.")
            ElseIf Character = "/"c AndAlso Index + 1 < Content.Length AndAlso (Content(Index + 1) = "/"c OrElse Content(Index + 1) = "*"c) Then
                If Not _Options.AllowJsonComments Then Throw New JsonReaderException("JSON comments are not allowed.")
                If Content(Index + 1) = "/"c Then
                    Index += 2
                    While Index < Content.Length AndAlso Content(Index) <> ChrW(10) AndAlso Content(Index) <> ChrW(13)
                        Index += 1
                    End While
                Else
                    Dim EndIndex As Integer = Content.IndexOf("*/", Index + 2, StringComparison.Ordinal)
                    If EndIndex < 0 Then Throw New JsonReaderException("Unterminated JSON comment.")
                    Index = EndIndex + 1
                End If
            ElseIf Not Char.IsWhiteSpace(Character) Then
                If Not _Options.AllowTrailingCommas AndAlso Previous = ","c AndAlso (Character = "}"c OrElse Character = "]"c) Then Throw New JsonReaderException("Trailing commas are not allowed.")
                Previous = Character
            End If
            Index += 1
        End While
    End Sub
    Private Function IsValidStoredValue(FilePath As String) As Boolean
        Try
            Dim IgnoredValue As T = DeserializeFromFile(FilePath)
            Return True
        Catch ex As Exception When IsRecoverablePrimaryFailure(ex)
            Return False
        End Try
    End Function
    Private Async Function IsValidStoredValueAsync(FilePath As String, CancellationToken As CancellationToken) As Task(Of Boolean)
        Try
            Dim IgnoredValue As T = Await DeserializeFromFileAsync(FilePath, CancellationToken).ConfigureAwait(False)
            Return True
        Catch ex As Exception When IsRecoverablePrimaryFailure(ex)
            Return False
        End Try
    End Function
    Private Shared Sub CopyFileAtomically(SourcePath As String, TargetPath As String)
        EnsureParentDirectory(TargetPath)
        Dim TemporaryPath As String = CreateTemporaryPath(TargetPath)
        Try
            Using InputStream As New FileStream(SourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.SequentialScan)
                Using OutputStream As New FileStream(TemporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.WriteThrough Or FileOptions.SequentialScan)
                    InputStream.CopyTo(OutputStream, FileBufferSize)
                    OutputStream.Flush(True)
                End Using
            End Using
            PromoteTemporaryFile(TemporaryPath, TargetPath)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Sub
    Private Shared Async Function CopyFileAtomicallyAsync(SourcePath As String, TargetPath As String, CancellationToken As CancellationToken) As Task
        EnsureParentDirectory(TargetPath)
        Dim TemporaryPath As String = CreateTemporaryPath(TargetPath)
        Try
            Using InputStream As New FileStream(SourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.Asynchronous Or FileOptions.SequentialScan)
                Using OutputStream As New FileStream(TemporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.Asynchronous Or FileOptions.WriteThrough Or FileOptions.SequentialScan)
                    Await InputStream.CopyToAsync(OutputStream, FileBufferSize, CancellationToken).ConfigureAwait(False)
                    Await OutputStream.FlushAsync(CancellationToken).ConfigureAwait(False)
                    OutputStream.Flush(True)
                End Using
            End Using
            PromoteTemporaryFile(TemporaryPath, TargetPath)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Function
    Private Shared Sub PromoteTemporaryFile(TemporaryPath As String, TargetPath As String)
        If File.Exists(TargetPath) Then
            ' Do not fall back to Delete+Move: that would break atomic replacement guarantees.
            File.Replace(TemporaryPath, TargetPath, Nothing, True)
        Else
            File.Move(TemporaryPath, TargetPath)
        End If
    End Sub
    Private Shared Sub EnsureParentDirectory(FilePath As String)
        Dim DirectoryPath As String = Path.GetDirectoryName(FilePath)
        If Not String.IsNullOrEmpty(DirectoryPath) Then Directory.CreateDirectory(DirectoryPath)
    End Sub
    Private Shared Function CreateTemporaryPath(TargetPath As String) As String
        Dim DirectoryPath As String = Path.GetDirectoryName(TargetPath)
        Dim FileName As String = Path.GetFileName(TargetPath)
        Return Path.Combine(DirectoryPath, $".{FileName}.{Guid.NewGuid():N}.tmp")
    End Function
    Private Shared Sub DeleteTemporaryFile(TemporaryPath As String)
        Try
            If File.Exists(TemporaryPath) Then File.Delete(TemporaryPath)
        Catch ex As IOException
        Catch ex As UnauthorizedAccessException
        End Try
    End Sub
    Private Shared Function IsRecoverablePrimaryFailure(ex As Exception) As Boolean
        Return TypeOf ex Is JsonException OrElse
               TypeOf ex Is DecoderFallbackException OrElse
               TypeOf ex Is FileNotFoundException OrElse
               TypeOf ex Is DirectoryNotFoundException
    End Function

    Private Shared Function IsExpectedLoadException(ex As Exception) As Boolean
        Return TypeOf ex Is IOException OrElse
               TypeOf ex Is UnauthorizedAccessException OrElse
               TypeOf ex Is DecoderFallbackException OrElse
               TypeOf ex Is JsonException
    End Function
End Class
