Imports System.IO
Imports System.Text
Imports CoreSuite.Services.Cryptography
Imports Microsoft.VisualStudio.TestTools.UnitTesting
<TestClass>
Public NotInheritable Class HashingTests
    Private Const EmptySha256 As String = "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855"
    Private Const AbcSha256 As String = "BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD"
    Private _TestFolder As String
    <TestInitialize>
    Public Sub Initialize()
        _TestFolder = Path.Combine(Path.GetTempPath(), "CoreSuite.Cryptography.Tests", Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_TestFolder)
    End Sub
    <TestCleanup>
    Public Sub Cleanup()
        If Directory.Exists(_TestFolder) Then Directory.Delete(_TestFolder, True)
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_EmptyText_ShouldMatchKnownHash()
        Dim ActualHash As String = Hashing.ComputeSha256(String.Empty)
        Assert.AreEqual(EmptySha256, ActualHash)
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_Text_ShouldMatchKnownHash()
        Dim ActualHash As String = Hashing.ComputeSha256("abc")
        Assert.AreEqual(AbcSha256, ActualHash)
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_Bytes_ShouldMatchKnownHash()
        Dim Data As Byte() = Encoding.UTF8.GetBytes("abc")
        Dim ActualHash As String = Hashing.ComputeSha256(Data)
        Assert.AreEqual(AbcSha256, ActualHash)
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_Stream_ShouldHashFromCurrentPosition()
        Dim Data As Byte() = Encoding.UTF8.GetBytes("prefixabc")
        Using DataStream As New MemoryStream(Data)
            DataStream.Position = Encoding.UTF8.GetByteCount("prefix")
            Dim ActualHash As String = Hashing.ComputeSha256(DataStream)
            Assert.AreEqual(AbcSha256, ActualHash)
        End Using
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_NullText_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() Hashing.ComputeSha256(DirectCast(Nothing, String)))
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_NullData_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() Hashing.ComputeSha256(DirectCast(Nothing, Byte())))
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_NullStream_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() Hashing.ComputeSha256(DirectCast(Nothing, Stream)))
    End Sub
    <TestMethod>
    Public Sub ComputeSha256_UnreadableStream_ShouldThrowArgumentException()
        Dim DataStream As New MemoryStream()
        DataStream.Dispose()
        Assert.ThrowsException(Of ArgumentException)(Sub() Hashing.ComputeSha256(DataStream))
    End Sub
    <TestMethod>
    Public Sub VerifySha256_MatchingText_ShouldReturnTrue()
        Assert.IsTrue(Hashing.VerifySha256("abc", AbcSha256))
    End Sub
    <TestMethod>
    Public Sub VerifySha256_DifferentText_ShouldReturnFalse()
        Assert.IsFalse(Hashing.VerifySha256("different", AbcSha256))
    End Sub
    <TestMethod>
    Public Sub VerifySha256_LowercaseHash_ShouldReturnTrue()
        Assert.IsTrue(Hashing.VerifySha256("abc", AbcSha256.ToLowerInvariant()))
    End Sub
    <TestMethod>
    Public Sub VerifySha256_InvalidExpectedHash_ShouldReturnFalse()
        Assert.IsFalse(Hashing.VerifySha256("abc", String.Empty))
        Assert.IsFalse(Hashing.VerifySha256("abc", "not-a-hash"))
        Assert.IsFalse(Hashing.VerifySha256("abc", "AA"))
    End Sub
    <TestMethod>
    Public Sub ComputeFileSha256_ValidFile_ShouldMatchKnownHash()
        Dim FilePath As String = CreateTestFile("sync.txt", "abc")
        Dim ActualHash As String = Hashing.ComputeFileSha256(FilePath)
        Assert.AreEqual(AbcSha256, ActualHash)
    End Sub
    <TestMethod>
    Public Async Function ComputeFileSha256Async_ValidFile_ShouldMatchKnownHash() As Task
        Dim FilePath As String = CreateTestFile("async.txt", "abc")
        Dim ActualHash As String = Await Hashing.ComputeFileSha256Async(FilePath)
        Assert.AreEqual(AbcSha256, ActualHash)
    End Function
    <TestMethod>
    Public Sub VerifyFileSha256_ValidAndInvalidHashes_ShouldReturnExpectedResults()
        Dim FilePath As String = CreateTestFile("verify.txt", "abc")
        Assert.IsTrue(Hashing.VerifyFileSha256(FilePath, AbcSha256))
        Assert.IsFalse(Hashing.VerifyFileSha256(FilePath, EmptySha256))
        Assert.IsFalse(Hashing.VerifyFileSha256(FilePath, "invalid"))
    End Sub
    <TestMethod>
    Public Async Function VerifyFileSha256Async_ValidAndInvalidHashes_ShouldReturnExpectedResults() As Task
        Dim FilePath As String = CreateTestFile("verify-async.txt", "abc")
        Assert.IsTrue(Await Hashing.VerifyFileSha256Async(FilePath, AbcSha256))
        Assert.IsFalse(Await Hashing.VerifyFileSha256Async(FilePath, EmptySha256))
    End Function
    <TestMethod>
    Public Sub ComputeFileSha256_EmptyPath_ShouldThrowArgumentException()
        Assert.ThrowsException(Of ArgumentException)(Sub() Hashing.ComputeFileSha256(String.Empty))
    End Sub
    <TestMethod>
    Public Sub ComputeFileSha256_MissingFile_ShouldThrowFileNotFoundException()
        Dim FilePath As String = Path.Combine(_TestFolder, "missing.txt")
        Assert.ThrowsException(Of FileNotFoundException)(Sub() Hashing.ComputeFileSha256(FilePath))
    End Sub
    Private Function CreateTestFile(FileName As String, Content As String) As String
        Dim FilePath As String = Path.Combine(_TestFolder, FileName)
        File.WriteAllText(FilePath, Content, New UTF8Encoding(False))
        Return FilePath
    End Function
End Class