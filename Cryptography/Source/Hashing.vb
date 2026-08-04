Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
''' <summary>
''' Provides SHA-256 hashing helpers for text, binary data, streams, and files.
''' </summary>
Public NotInheritable Class Hashing
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Computes the SHA-256 hash of UTF-8 text and returns an uppercase hexadecimal string.
    ''' </summary>
    ''' <param name="Text">The text to hash.</param>
    ''' <returns>The SHA-256 hash as 64 uppercase hexadecimal characters.</returns>
    Public Shared Function ComputeSha256(Text As String) As String
        ArgumentNullException.ThrowIfNull(Text)
        Dim DataBytes As Byte() = Encoding.UTF8.GetBytes(Text)
        Try
            Return ComputeSha256(DataBytes)
        Finally
            CryptographicOperations.ZeroMemory(DataBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Computes the SHA-256 hash of binary data and returns an uppercase hexadecimal string.
    ''' </summary>
    ''' <param name="Data">The data to hash.</param>
    ''' <returns>The SHA-256 hash as 64 uppercase hexadecimal characters.</returns>
    Public Shared Function ComputeSha256(Data As Byte()) As String
        ArgumentNullException.ThrowIfNull(Data)
        Return Convert.ToHexString(SHA256.HashData(Data))
    End Function
    ''' <summary>
    ''' Computes the SHA-256 hash of a stream from its current position and returns an uppercase hexadecimal string.
    ''' </summary>
    ''' <param name="Stream">The readable stream to hash.</param>
    ''' <returns>The SHA-256 hash as 64 uppercase hexadecimal characters.</returns>
    Public Shared Function ComputeSha256(Stream As Stream) As String
        ArgumentNullException.ThrowIfNull(Stream)
        If Not Stream.CanRead Then Throw New ArgumentException("The stream must be readable.", NameOf(Stream))
        Return Convert.ToHexString(SHA256.HashData(Stream))
    End Function
    ''' <summary>
    ''' Computes the SHA-256 hash of a file and returns an uppercase hexadecimal string.
    ''' </summary>
    ''' <param name="FilePath">The path of the file to hash.</param>
    ''' <returns>The SHA-256 hash as 64 uppercase hexadecimal characters.</returns>
    Public Shared Function ComputeFileSha256(FilePath As String) As String
        ValidateFilePath(FilePath)
        Using Stream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81_920, FileOptions.SequentialScan)
            Return ComputeSha256(Stream)
        End Using
    End Function
    ''' <summary>
    ''' Asynchronously computes the SHA-256 hash of a file and returns an uppercase hexadecimal string.
    ''' </summary>
    ''' <param name="FilePath">The path of the file to hash.</param>
    ''' <param name="CancellationToken">The token used to cancel the operation.</param>
    ''' <returns>The SHA-256 hash as 64 uppercase hexadecimal characters.</returns>
    Public Shared Async Function ComputeFileSha256Async(FilePath As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of String)
        ValidateFilePath(FilePath)
        Using Stream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81_920, FileOptions.Asynchronous Or FileOptions.SequentialScan)
            Dim HashBytes As Byte() = Await SHA256.HashDataAsync(Stream, CancellationToken)
            Return Convert.ToHexString(HashBytes)
        End Using
    End Function
    ''' <summary>
    ''' Verifies whether UTF-8 text matches an expected SHA-256 hexadecimal hash using a fixed-time byte comparison.
    ''' </summary>
    ''' <param name="Text">The text to verify.</param>
    ''' <param name="ExpectedHash">The expected 64-character hexadecimal SHA-256 hash.</param>
    ''' <returns><see langword="True"/> when the hash matches; otherwise, <see langword="False"/>.</returns>
    Public Shared Function VerifySha256(Text As String, ExpectedHash As String) As Boolean
        ArgumentNullException.ThrowIfNull(Text)
        Dim DataBytes As Byte() = Encoding.UTF8.GetBytes(Text)
        Try
            Return VerifySha256(DataBytes, ExpectedHash)
        Finally
            CryptographicOperations.ZeroMemory(DataBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Verifies whether binary data matches an expected SHA-256 hexadecimal hash using a fixed-time byte comparison.
    ''' </summary>
    ''' <param name="Data">The data to verify.</param>
    ''' <param name="ExpectedHash">The expected 64-character hexadecimal SHA-256 hash.</param>
    ''' <returns><see langword="True"/> when the hash matches; otherwise, <see langword="False"/>.</returns>
    Public Shared Function VerifySha256(Data As Byte(), ExpectedHash As String) As Boolean
        ArgumentNullException.ThrowIfNull(Data)
        If String.IsNullOrWhiteSpace(ExpectedHash) Then Return False
        Dim ExpectedBytes As Byte()
        Try
            ExpectedBytes = Convert.FromHexString(ExpectedHash)
        Catch Ex As FormatException
            Return False
        End Try
        If ExpectedBytes.Length <> 32 Then Return False
        Dim ComputedBytes As Byte() = SHA256.HashData(Data)
        Try
            Return CryptographicOperations.FixedTimeEquals(ComputedBytes, ExpectedBytes)
        Finally
            CryptographicOperations.ZeroMemory(ComputedBytes.AsSpan())
            CryptographicOperations.ZeroMemory(ExpectedBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Verifies whether a file matches an expected SHA-256 hexadecimal hash.
    ''' </summary>
    ''' <param name="FilePath">The path of the file to verify.</param>
    ''' <param name="ExpectedHash">The expected 64-character hexadecimal SHA-256 hash.</param>
    ''' <returns><see langword="True"/> when the hash matches; otherwise, <see langword="False"/>.</returns>
    Public Shared Function VerifyFileSha256(FilePath As String, ExpectedHash As String) As Boolean
        Dim ActualHash As String = ComputeFileSha256(FilePath)
        Return FixedTimeHexEquals(ActualHash, ExpectedHash)
    End Function
    ''' <summary>
    ''' Asynchronously verifies whether a file matches an expected SHA-256 hexadecimal hash.
    ''' </summary>
    ''' <param name="FilePath">The path of the file to verify.</param>
    ''' <param name="ExpectedHash">The expected 64-character hexadecimal SHA-256 hash.</param>
    ''' <param name="CancellationToken">The token used to cancel the operation.</param>
    ''' <returns><see langword="True"/> when the hash matches; otherwise, <see langword="False"/>.</returns>
    Public Shared Async Function VerifyFileSha256Async(FilePath As String, ExpectedHash As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        Dim ActualHash As String = Await ComputeFileSha256Async(FilePath, CancellationToken)
        Return FixedTimeHexEquals(ActualHash, ExpectedHash)
    End Function
    Private Shared Function FixedTimeHexEquals(ActualHash As String, ExpectedHash As String) As Boolean
        If String.IsNullOrWhiteSpace(ExpectedHash) Then Return False
        Dim ActualBytes As Byte()
        Dim ExpectedBytes As Byte()
        Try
            ActualBytes = Convert.FromHexString(ActualHash)
            ExpectedBytes = Convert.FromHexString(ExpectedHash)
        Catch Ex As FormatException
            Return False
        End Try
        If ActualBytes.Length <> ExpectedBytes.Length Then Return False
        Try
            Return CryptographicOperations.FixedTimeEquals(ActualBytes, ExpectedBytes)
        Finally
            CryptographicOperations.ZeroMemory(ActualBytes.AsSpan())
            CryptographicOperations.ZeroMemory(ExpectedBytes.AsSpan())
        End Try
    End Function
    Private Shared Sub ValidateFilePath(FilePath As String)
        If String.IsNullOrWhiteSpace(FilePath) Then Throw New ArgumentException("File path cannot be empty.", NameOf(FilePath))
    End Sub
End Class
