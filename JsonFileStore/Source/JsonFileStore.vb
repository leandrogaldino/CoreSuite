Imports System.IO
Imports System.Text.Json
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
    Private _AutoRecoverFromBackup As Boolean = True
    ''' <summary>
    ''' Initializes a new instance of the <see cref="JsonFileStore(Of T)"/> class.
    ''' </summary>
    ''' <param name="FilePath">The path of the primary JSON file.</param>
    ''' <param name="SerializerOptions">Optional JSON serializer settings. The settings are cloned by the store.</param>
    ''' <param name="BackupPath">An optional backup path. The default is the primary path followed by <c>.bak</c>.</param>
    ''' <exception cref="ArgumentNullException"><paramref name="FilePath"/> is <see langword="Nothing"/>.</exception>
    ''' <exception cref="ArgumentException">A path is empty, contains only whitespace, or both paths resolve to the same file.</exception>
    Public Sub New(FilePath As String, Optional SerializerOptions As JsonSerializerOptions = Nothing, Optional BackupPath As String = Nothing)
        ArgumentNullException.ThrowIfNull(FilePath)
        If String.IsNullOrWhiteSpace(FilePath) Then Throw New ArgumentException("The JSON file path cannot be empty or whitespace.", NameOf(FilePath))
        _FilePath = Path.GetFullPath(FilePath)
        If BackupPath Is Nothing Then
            _BackupPath = _FilePath & ".bak"
        Else
            If String.IsNullOrWhiteSpace(BackupPath) Then Throw New ArgumentException("The backup file path cannot be empty or whitespace.", NameOf(BackupPath))
            _BackupPath = Path.GetFullPath(BackupPath)
        End If
        Dim pathComparison As StringComparison = If(OperatingSystem.IsWindows(), StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
        If String.Equals(_FilePath, _BackupPath, pathComparison) Then Throw New ArgumentException("The primary and backup paths must be different.", NameOf(BackupPath))
        SerializerOptions = If(SerializerOptions Is Nothing, New JsonSerializerOptions(), New JsonSerializerOptions(SerializerOptions))
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
    ''' Gets the serializer settings used by this store.
    ''' </summary>
    ''' <value>A private clone of the settings supplied to the constructor.</value>
    ''' <remarks>Configure these settings before the first operation. System.Text.Json makes options read-only after use.</remarks>
    Public ReadOnly Property SerializerOptions As JsonSerializerOptions
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
            If File.Exists(_FilePath) AndAlso IsValidStoredValue(_FilePath) Then CopyFileAtomically(_FilePath, _BackupPath)
            PromoteTemporaryFile(TemporaryPath, _FilePath)
            If Not File.Exists(_BackupPath) Then CopyFileAtomically(_FilePath, _BackupPath)
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
            If File.Exists(_FilePath) AndAlso Await IsValidStoredValueAsync(_FilePath, CancellationToken).ConfigureAwait(False) Then Await CopyFileAtomicallyAsync(_FilePath, _BackupPath, CancellationToken).ConfigureAwait(False)
            PromoteTemporaryFile(TemporaryPath, _FilePath)
            If Not File.Exists(_BackupPath) Then Await CopyFileAtomicallyAsync(_FilePath, _BackupPath, CancellationToken.None).ConfigureAwait(False)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Function
    Private Function LoadCore() As T
        Dim PrimaryException As Exception
        Try
            Return DeserializeFromFile(_FilePath)
        Catch ex As Exception When IsExpectedLoadException(ex)
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
        Catch ex As Exception When IsExpectedLoadException(ex)
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
            JsonSerializer.Serialize(OutputStream, Value, SerializerOptions)
            OutputStream.Flush(True)
        End Using
    End Sub
    Private Async Function SerializeToFileAsync(FilePath As String, Value As T, CancellationToken As CancellationToken) As Task
        Using OutputStream As New FileStream(FilePath, FileMode.CreateNew, FileAccess.Write, FileShare.None, FileBufferSize, FileOptions.Asynchronous Or FileOptions.WriteThrough Or FileOptions.SequentialScan)
            Await JsonSerializer.SerializeAsync(OutputStream, Value, SerializerOptions, CancellationToken).ConfigureAwait(False)
            Await OutputStream.FlushAsync(CancellationToken).ConfigureAwait(False)
            OutputStream.Flush(True)
        End Using
    End Function
    Private Function DeserializeFromFile(FilePath As String) As T
        Using InputStream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.SequentialScan)
            Return JsonSerializer.Deserialize(Of T)(InputStream, SerializerOptions)
        End Using
    End Function
    Private Async Function DeserializeFromFileAsync(FilePath As String, CancellationToken As CancellationToken) As Task(Of T)
        Using InputStream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, FileBufferSize, FileOptions.Asynchronous Or FileOptions.SequentialScan)
            Return Await JsonSerializer.DeserializeAsync(Of T)(InputStream, SerializerOptions, CancellationToken).ConfigureAwait(False)
        End Using
    End Function
    Private Function IsValidStoredValue(FilePath As String) As Boolean
        Try
            Dim IgnoredValue As T = DeserializeFromFile(FilePath)
            Return True
        Catch ex As Exception When IsExpectedLoadException(ex)
            Return False
        End Try
    End Function
    Private Async Function IsValidStoredValueAsync(FilePath As String, CancellationToken As CancellationToken) As Task(Of Boolean)
        Try
            Dim IgnoredValue As T = Await DeserializeFromFileAsync(FilePath, CancellationToken).ConfigureAwait(False)
            Return True
        Catch ex As Exception When IsExpectedLoadException(ex)
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
            Try
                File.Replace(TemporaryPath, TargetPath, Nothing, True)
            Catch ex As PlatformNotSupportedException
                File.Move(TemporaryPath, TargetPath, True)
            End Try
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
    Private Shared Function IsExpectedLoadException(Exception As Exception) As Boolean
        Return TypeOf Exception Is IOException OrElse TypeOf Exception Is UnauthorizedAccessException OrElse TypeOf Exception Is JsonException OrElse TypeOf Exception Is NotSupportedException
    End Function
End Class
