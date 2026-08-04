Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Threading
''' <summary>
''' Stores JSON or binary data in a versioned file protected by Windows Data Protection API.
''' </summary>
Public NotInheritable Class ProtectedFileStorage
    Public Const CurrentFormatVersion As Byte = 1
    Private Const MagicSize As Integer = 4
    Private Const HeaderSize As Integer = 10
    Private Shared ReadOnly MagicBytes As Byte() = {&H43, &H53, &H44, &H50}
    Private Shared ReadOnly DefaultJsonOptions As New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True, .WriteIndented = True}
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Serializes a value as JSON and stores it in a DPAPI-protected file.
    ''' </summary>
    ''' <typeparam name="T">The type of value to store.</typeparam>
    ''' <param name="FilePath">The destination file path.</param>
    ''' <param name="Value">The value to serialize and protect.</param>
    ''' <param name="Scope">The Windows DPAPI protection scope.</param>
    ''' <param name="Entropy">Optional additional entropy that must also be supplied when loading the file.</param>
    ''' <param name="Options">Optional JSON serialization options.</param>
    Public Shared Sub Save(Of T)(FilePath As String, Value As T, Optional Scope As DataProtectionScope = DataProtectionScope.CurrentUser, Optional Entropy As Byte() = Nothing, Optional Options As JsonSerializerOptions = Nothing)
        Dim JsonBytes As Byte() = JsonSerializer.SerializeToUtf8Bytes(Value, ResolveOptions(Options))
        Try
            SaveBytes(FilePath, JsonBytes, Scope, Entropy)
        Finally
            CryptographicOperations.ZeroMemory(JsonBytes.AsSpan())
        End Try
    End Sub
    ''' <summary>
    ''' Asynchronously serializes a value as JSON and stores it in a DPAPI-protected file.
    ''' </summary>
    ''' <typeparam name="T">The type of value to store.</typeparam>
    ''' <param name="FilePath">The destination file path.</param>
    ''' <param name="Value">The value to serialize and protect.</param>
    ''' <param name="Scope">The Windows DPAPI protection scope.</param>
    ''' <param name="Entropy">Optional additional entropy that must also be supplied when loading the file.</param>
    ''' <param name="Options">Optional JSON serialization options.</param>
    ''' <param name="CancellationToken">The token used to cancel the file operation.</param>
    Public Shared Async Function SaveAsync(Of T)(FilePath As String, Value As T, Optional Scope As DataProtectionScope = DataProtectionScope.CurrentUser, Optional Entropy As Byte() = Nothing, Optional Options As JsonSerializerOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        Dim JsonBytes As Byte() = JsonSerializer.SerializeToUtf8Bytes(Value, ResolveOptions(Options))
        Try
            Await SaveBytesAsync(FilePath, JsonBytes, Scope, Entropy, CancellationToken)
        Finally
            CryptographicOperations.ZeroMemory(JsonBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Loads, unprotects, and deserializes a JSON value from a protected file.
    ''' </summary>
    ''' <typeparam name="T">The type to deserialize.</typeparam>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <param name="Options">Optional JSON serialization options.</param>
    ''' <returns>The deserialized value.</returns>
    Public Shared Function Load(Of T)(FilePath As String, Optional Entropy As Byte() = Nothing, Optional Options As JsonSerializerOptions = Nothing) As T
        Dim JsonBytes As Byte() = LoadBytes(FilePath, Entropy)
        Try
            Return JsonSerializer.Deserialize(Of T)(JsonBytes, ResolveOptions(Options))
        Finally
            CryptographicOperations.ZeroMemory(JsonBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Asynchronously loads, unprotects, and deserializes a JSON value from a protected file.
    ''' </summary>
    ''' <typeparam name="T">The type to deserialize.</typeparam>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <param name="Options">Optional JSON serialization options.</param>
    ''' <param name="CancellationToken">The token used to cancel the file operation.</param>
    ''' <returns>The deserialized value.</returns>
    Public Shared Async Function LoadAsync(Of T)(FilePath As String, Optional Entropy As Byte() = Nothing, Optional Options As JsonSerializerOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of T)
        Dim JsonBytes As Byte() = Await LoadBytesAsync(FilePath, Entropy, CancellationToken)
        Try
            Return JsonSerializer.Deserialize(Of T)(JsonBytes, ResolveOptions(Options))
        Finally
            CryptographicOperations.ZeroMemory(JsonBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Attempts to load a JSON value from a protected file without throwing for missing, malformed, inaccessible, or undecryptable files.
    ''' </summary>
    ''' <typeparam name="T">The type to deserialize.</typeparam>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="Value">Receives the deserialized value when successful.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <param name="Options">Optional JSON serialization options.</param>
    ''' <returns><see langword="True"/> when loading succeeds; otherwise, <see langword="False"/>.</returns>
    Public Shared Function TryLoad(Of T)(FilePath As String, ByRef Value As T, Optional Entropy As Byte() = Nothing, Optional Options As JsonSerializerOptions = Nothing) As Boolean
        Value = Nothing
        Try
            Value = Load(Of T)(FilePath, Entropy, Options)
            Return True
        Catch Ex As PlatformNotSupportedException
            Throw
        Catch Ex As FileNotFoundException
            Return False
        Catch Ex As DirectoryNotFoundException
            Return False
        Catch Ex As UnauthorizedAccessException
            Return False
        Catch Ex As IOException
            Return False
        Catch Ex As CryptographicException
            Return False
        Catch Ex As JsonException
            Return False
        Catch Ex As FormatException
            Return False
        Catch Ex As NotSupportedException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Loads a protected JSON value or returns a caller-provided default value when loading fails.
    ''' </summary>
    ''' <typeparam name="T">The type to deserialize.</typeparam>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="DefaultValue">The value returned when loading fails.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <param name="Options">Optional JSON serialization options.</param>
    ''' <returns>The loaded value or the provided default value.</returns>
    Public Shared Function LoadOrDefault(Of T)(FilePath As String, DefaultValue As T, Optional Entropy As Byte() = Nothing, Optional Options As JsonSerializerOptions = Nothing) As T
        Dim Value As T = Nothing
        If TryLoad(FilePath, Value, Entropy, Options) Then Return Value
        Return DefaultValue
    End Function
    ''' <summary>
    ''' Protects and stores arbitrary binary data in a versioned file.
    ''' </summary>
    ''' <param name="FilePath">The destination file path.</param>
    ''' <param name="Data">The binary data to protect.</param>
    ''' <param name="Scope">The Windows DPAPI protection scope.</param>
    ''' <param name="Entropy">Optional additional entropy that must also be supplied when loading the file.</param>
    Public Shared Sub SaveBytes(FilePath As String, Data As Byte(), Optional Scope As DataProtectionScope = DataProtectionScope.CurrentUser, Optional Entropy As Byte() = Nothing)
        ValidatePlatform()
        ValidateFilePath(FilePath)
        ArgumentNullException.ThrowIfNull(Data)
        ValidateScope(Scope)
        Dim ProtectedBytes As Byte() = ProtectedData.Protect(Data, Entropy, Scope)
        Try
            Dim Package As Byte() = CreatePackage(ProtectedBytes, Scope)
            WriteFileAtomically(FilePath, Package)
        Finally
            CryptographicOperations.ZeroMemory(ProtectedBytes.AsSpan())
        End Try
    End Sub
    ''' <summary>
    ''' Asynchronously protects and stores arbitrary binary data in a versioned file.
    ''' </summary>
    ''' <param name="FilePath">The destination file path.</param>
    ''' <param name="Data">The binary data to protect.</param>
    ''' <param name="Scope">The Windows DPAPI protection scope.</param>
    ''' <param name="Entropy">Optional additional entropy that must also be supplied when loading the file.</param>
    ''' <param name="CancellationToken">The token used to cancel the file operation.</param>
    Public Shared Async Function SaveBytesAsync(FilePath As String, Data As Byte(), Optional Scope As DataProtectionScope = DataProtectionScope.CurrentUser, Optional Entropy As Byte() = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        ValidatePlatform()
        ValidateFilePath(FilePath)
        ArgumentNullException.ThrowIfNull(Data)
        ValidateScope(Scope)
        CancellationToken.ThrowIfCancellationRequested()
        Dim ProtectedBytes As Byte() = ProtectedData.Protect(Data, Entropy, Scope)
        Try
            Dim Package As Byte() = CreatePackage(ProtectedBytes, Scope)
            Await WriteFileAtomicallyAsync(FilePath, Package, CancellationToken)
        Finally
            CryptographicOperations.ZeroMemory(ProtectedBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Loads and unprotects arbitrary binary data from a protected file.
    ''' </summary>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <returns>The unprotected binary data.</returns>
    Public Shared Function LoadBytes(FilePath As String, Optional Entropy As Byte() = Nothing) As Byte()
        ValidatePlatform()
        ValidateFilePath(FilePath)
        Dim Package As Byte() = File.ReadAllBytes(FilePath)
        Dim Scope As DataProtectionScope
        Dim ProtectedBytes As Byte() = ParsePackage(Package, Scope)
        Try
            Return ProtectedData.Unprotect(ProtectedBytes, Entropy, Scope)
        Finally
            CryptographicOperations.ZeroMemory(ProtectedBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Asynchronously loads and unprotects arbitrary binary data from a protected file.
    ''' </summary>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <param name="CancellationToken">The token used to cancel the file operation.</param>
    ''' <returns>The unprotected binary data.</returns>
    Public Shared Async Function LoadBytesAsync(FilePath As String, Optional Entropy As Byte() = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Byte())
        ValidatePlatform()
        ValidateFilePath(FilePath)
        Dim Package As Byte() = Await File.ReadAllBytesAsync(FilePath, CancellationToken)
        Dim Scope As DataProtectionScope
        Dim ProtectedBytes As Byte() = ParsePackage(Package, Scope)
        Try
            Return ProtectedData.Unprotect(ProtectedBytes, Entropy, Scope)
        Finally
            CryptographicOperations.ZeroMemory(ProtectedBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Attempts to load arbitrary binary data from a protected file.
    ''' </summary>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <param name="Data">Receives the unprotected data when successful.</param>
    ''' <param name="Entropy">The same optional additional entropy used when saving.</param>
    ''' <returns><see langword="True"/> when loading succeeds; otherwise, <see langword="False"/>.</returns>
    Public Shared Function TryLoadBytes(FilePath As String, ByRef Data As Byte(), Optional Entropy As Byte() = Nothing) As Boolean
        Data = Array.Empty(Of Byte)()
        Try
            Data = LoadBytes(FilePath, Entropy)
            Return True
        Catch Ex As PlatformNotSupportedException
            Throw
        Catch Ex As FileNotFoundException
            Return False
        Catch Ex As DirectoryNotFoundException
            Return False
        Catch Ex As UnauthorizedAccessException
            Return False
        Catch Ex As IOException
            Return False
        Catch Ex As CryptographicException
            Return False
        Catch Ex As FormatException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Determines whether a file exists and contains a structurally valid package created by this class.
    ''' </summary>
    ''' <param name="FilePath">The file path to inspect.</param>
    ''' <returns><see langword="True"/> when the package structure is valid; otherwise, <see langword="False"/>.</returns>
    Public Shared Function IsProtectedFile(FilePath As String) As Boolean
        If String.IsNullOrWhiteSpace(FilePath) OrElse Not File.Exists(FilePath) Then Return False
        Try
            Dim Package As Byte() = File.ReadAllBytes(FilePath)
            Dim Scope As DataProtectionScope
            ParsePackage(Package, Scope)
            Return True
        Catch Ex As IOException
            Return False
        Catch Ex As UnauthorizedAccessException
            Return False
        Catch Ex As FormatException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Gets the DPAPI scope stored in a protected file header without decrypting its content.
    ''' </summary>
    ''' <param name="FilePath">The protected file path.</param>
    ''' <returns>The stored DPAPI protection scope.</returns>
    Public Shared Function GetProtectionScope(FilePath As String) As DataProtectionScope
        ValidateFilePath(FilePath)
        Dim Package As Byte() = File.ReadAllBytes(FilePath)
        Dim Scope As DataProtectionScope
        ParsePackage(Package, Scope)
        Return Scope
    End Function
    ''' <summary>
    ''' Deletes a protected file when it exists.
    ''' </summary>
    ''' <param name="FilePath">The file path to delete.</param>
    ''' <returns><see langword="True"/> when a file existed and was deleted; otherwise, <see langword="False"/>.</returns>
    Public Shared Function Delete(FilePath As String) As Boolean
        ValidateFilePath(FilePath)
        If Not File.Exists(FilePath) Then Return False
        File.Delete(FilePath)
        Return True
    End Function
    ''' <summary>
    ''' Creates a new set of JSON options equivalent to the defaults used by this class.
    ''' </summary>
    ''' <returns>A mutable JSON serializer options instance.</returns>
    Public Shared Function CreateDefaultJsonOptions() As JsonSerializerOptions
        Return New JsonSerializerOptions(DefaultJsonOptions)
    End Function
    Private Shared Function CreatePackage(ProtectedBytes As Byte(), Scope As DataProtectionScope) As Byte()
        Dim Header As Byte()
        Using Stream As New MemoryStream(HeaderSize)
            Using Writer As New BinaryWriter(Stream, Encoding.UTF8, True)
                Writer.Write(MagicBytes)
                Writer.Write(CurrentFormatVersion)
                Writer.Write(CByte(Scope))
                Writer.Write(ProtectedBytes.Length)
            End Using
            Header = Stream.ToArray()
        End Using
        If Header.Length <> HeaderSize Then Throw New InvalidOperationException("The protected storage header has an unexpected size.")
        Dim Package(Header.Length + ProtectedBytes.Length - 1) As Byte
        Buffer.BlockCopy(Header, 0, Package, 0, Header.Length)
        If ProtectedBytes.Length > 0 Then Buffer.BlockCopy(ProtectedBytes, 0, Package, Header.Length, ProtectedBytes.Length)
        Return Package
    End Function
    Private Shared Function ParsePackage(Package As Byte(), ByRef Scope As DataProtectionScope) As Byte()
        ArgumentNullException.ThrowIfNull(Package)
        If Package.Length < HeaderSize Then Throw New FormatException("The protected storage file is incomplete.")
        Dim ProtectedLength As Integer
        Using Stream As New MemoryStream(Package, False)
            Using Reader As New BinaryReader(Stream, Encoding.UTF8, True)
                Dim StoredMagic As Byte() = Reader.ReadBytes(MagicSize)
                If Not ByteArraysEqual(StoredMagic, MagicBytes) Then Throw New FormatException("The protected storage file identifier is invalid.")
                Dim Version As Byte = Reader.ReadByte()
                If Version <> CurrentFormatVersion Then Throw New FormatException($"Protected storage format version {Version} is not supported.")
                Scope = CType(Reader.ReadByte(), DataProtectionScope)
                ValidateScopeForFormat(Scope)
                ProtectedLength = Reader.ReadInt32()
            End Using
        End Using
        If ProtectedLength <= 0 Then Throw New FormatException("The protected storage payload length is invalid.")
        Dim ExpectedLength As Long = CLng(HeaderSize) + ProtectedLength
        If ExpectedLength <> Package.LongLength Then Throw New FormatException("The protected storage file length is invalid.")
        Dim ProtectedBytes(ProtectedLength - 1) As Byte
        Buffer.BlockCopy(Package, HeaderSize, ProtectedBytes, 0, ProtectedLength)
        Return ProtectedBytes
    End Function
    Private Shared Sub WriteFileAtomically(FilePath As String, Data As Byte())
        Dim FullPath As String = PrepareDestination(FilePath)
        Dim Folder As String = Path.GetDirectoryName(FullPath)
        Dim TempPath As String = Path.Combine(Folder, $".{Path.GetFileName(FullPath)}.{Guid.NewGuid():N}.tmp")
        Try
            Using Stream As New FileStream(TempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81_920, FileOptions.WriteThrough)
                Stream.Write(Data, 0, Data.Length)
                Stream.Flush(True)
            End Using
            File.Move(TempPath, FullPath, True)
        Finally
            If File.Exists(TempPath) Then File.Delete(TempPath)
        End Try
    End Sub
    Private Shared Async Function WriteFileAtomicallyAsync(FilePath As String, Data As Byte(), CancellationToken As CancellationToken) As Task
        Dim FullPath As String = PrepareDestination(FilePath)
        Dim Folder As String = Path.GetDirectoryName(FullPath)
        Dim TempPath As String = Path.Combine(Folder, $".{Path.GetFileName(FullPath)}.{Guid.NewGuid():N}.tmp")
        Try
            Using Stream As New FileStream(TempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81_920, FileOptions.Asynchronous Or FileOptions.WriteThrough)
                Await Stream.WriteAsync(Data.AsMemory(), CancellationToken)
                Await Stream.FlushAsync(CancellationToken)
                Stream.Flush(True)
            End Using
            CancellationToken.ThrowIfCancellationRequested()
            File.Move(TempPath, FullPath, True)
        Finally
            If File.Exists(TempPath) Then File.Delete(TempPath)
        End Try
    End Function
    Private Shared Function PrepareDestination(FilePath As String) As String
        ValidateFilePath(FilePath)
        Dim FullPath As String = Path.GetFullPath(FilePath)
        Dim Folder As String = Path.GetDirectoryName(FullPath)
        If String.IsNullOrWhiteSpace(Folder) Then Throw New ArgumentException("The file path must contain a valid directory.", NameOf(FilePath))
        Directory.CreateDirectory(Folder)
        Return FullPath
    End Function
    Private Shared Function ResolveOptions(Options As JsonSerializerOptions) As JsonSerializerOptions
        Return If(Options, DefaultJsonOptions)
    End Function
    Private Shared Sub ValidatePlatform()
        If Not OperatingSystem.IsWindows() Then Throw New PlatformNotSupportedException("ProtectedFileStorage requires Windows because it uses DPAPI.")
    End Sub
    Private Shared Sub ValidateFilePath(FilePath As String)
        If String.IsNullOrWhiteSpace(FilePath) Then Throw New ArgumentException("File path cannot be empty.", NameOf(FilePath))
    End Sub
    Private Shared Sub ValidateScope(Scope As DataProtectionScope)
        If Scope <> DataProtectionScope.CurrentUser AndAlso Scope <> DataProtectionScope.LocalMachine Then Throw New ArgumentOutOfRangeException(NameOf(Scope))
    End Sub
    Private Shared Sub ValidateScopeForFormat(Scope As DataProtectionScope)
        If Scope <> DataProtectionScope.CurrentUser AndAlso Scope <> DataProtectionScope.LocalMachine Then Throw New FormatException("The protected storage file contains an invalid protection scope.")
    End Sub
    Private Shared Function ByteArraysEqual(First As Byte(), Second As Byte()) As Boolean
        If First.Length <> Second.Length Then Return False
        For Index As Integer = 0 To First.Length - 1
            If First(Index) <> Second(Index) Then Return False
        Next
        Return True
    End Function
End Class
