Imports System.Security.Cryptography
Imports CoreSuite.Services.Cryptography
Imports Microsoft.VisualStudio.TestTools.UnitTesting
<TestClass>
Public NotInheritable Class PasswordHasherTests
    Private Const TestIterations As Integer = PasswordHasher.MinimumIterations
    Private Const TestPassword As String = "CoreSuite-Test-Password-2026!"
    <TestMethod>
    Public Sub HashPassword_ValidPassword_ShouldCreateSupportedSelfContainedFormat()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim Parts As String() = EncodedHash.Split("$"c)
        Assert.AreEqual(6, Parts.Length)
        Assert.AreEqual("CSPH", Parts(0))
        Assert.AreEqual(PasswordHasher.CurrentFormatVersion.ToString(), Parts(1))
        Assert.AreEqual("PBKDF2-SHA256", Parts(2))
        Assert.AreEqual(TestIterations.ToString(), Parts(3))
        Assert.AreEqual(16, Convert.FromBase64String(Parts(4)).Length)
        Assert.AreEqual(32, Convert.FromBase64String(Parts(5)).Length)
    End Sub
    <TestMethod>
    Public Sub HashPassword_SamePasswordTwice_ShouldGenerateDifferentHashes()
        Dim FirstHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim SecondHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Assert.AreNotEqual(FirstHash, SecondHash)
        Assert.IsTrue(PasswordHasher.VerifyPassword(TestPassword, FirstHash))
        Assert.IsTrue(PasswordHasher.VerifyPassword(TestPassword, SecondHash))
    End Sub
    <TestMethod>
    Public Sub VerifyPassword_CorrectPassword_ShouldReturnTrue()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Assert.IsTrue(PasswordHasher.VerifyPassword(TestPassword, EncodedHash))
    End Sub
    <TestMethod>
    Public Sub VerifyPassword_WrongPassword_ShouldReturnFalse()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Assert.IsFalse(PasswordHasher.VerifyPassword("Wrong password", EncodedHash))
    End Sub
    <TestMethod>
    Public Sub VerifyPasswordDetailed_CurrentConfiguration_ShouldReturnSuccess()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim Result As PasswordVerificationResult = PasswordHasher.VerifyPasswordDetailed(TestPassword, EncodedHash, TestIterations)
        Assert.AreEqual(PasswordVerificationResult.Success, Result)
    End Sub
    <TestMethod>
    Public Sub VerifyPasswordDetailed_LowerIterationCount_ShouldReturnSuccessRehashNeeded()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim Result As PasswordVerificationResult = PasswordHasher.VerifyPasswordDetailed(TestPassword, EncodedHash, TestIterations + 1)
        Assert.AreEqual(PasswordVerificationResult.SuccessRehashNeeded, Result)
    End Sub
    <TestMethod>
    Public Sub VerifyPasswordDetailed_WrongPassword_ShouldReturnFailed()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim Result As PasswordVerificationResult = PasswordHasher.VerifyPasswordDetailed("Wrong password", EncodedHash, TestIterations)
        Assert.AreEqual(PasswordVerificationResult.Failed, Result)
    End Sub
    <TestMethod>
    Public Sub VerifyPasswordDetailed_MalformedHash_ShouldReturnFailed()
        Dim Result As PasswordVerificationResult = PasswordHasher.VerifyPasswordDetailed(TestPassword, "invalid-hash", TestIterations)
        Assert.AreEqual(PasswordVerificationResult.Failed, Result)
    End Sub
    <TestMethod>
    Public Sub VerifyPassword_TamperedHash_ShouldReturnFalse()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim Parts As String() = EncodedHash.Split("$"c)
        Dim HashBytes As Byte() = Convert.FromBase64String(Parts(5))
        HashBytes(0) = HashBytes(0) Xor CByte(1)
        Parts(5) = Convert.ToBase64String(HashBytes)
        Dim TamperedHash As String = String.Join("$", Parts)
        Assert.IsFalse(PasswordHasher.VerifyPassword(TestPassword, TamperedHash))
    End Sub
    <TestMethod>
    Public Sub NeedsRehash_CurrentConfiguration_ShouldReturnFalse()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Assert.IsFalse(PasswordHasher.NeedsRehash(EncodedHash, TestIterations))
    End Sub
    <TestMethod>
    Public Sub NeedsRehash_HigherRequiredIterations_ShouldReturnTrue()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Assert.IsTrue(PasswordHasher.NeedsRehash(EncodedHash, TestIterations + 1))
    End Sub
    <TestMethod>
    Public Sub NeedsRehash_MalformedHash_ShouldReturnTrue()
        Assert.IsTrue(PasswordHasher.NeedsRehash("invalid-hash", TestIterations))
    End Sub
    <TestMethod>
    Public Sub TryGetIterations_ValidHash_ShouldReturnStoredIterations()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Dim Iterations As Integer
        Dim Result As Boolean = PasswordHasher.TryGetIterations(EncodedHash, Iterations)
        Assert.IsTrue(Result)
        Assert.AreEqual(TestIterations, Iterations)
    End Sub
    <TestMethod>
    Public Sub TryGetIterations_MalformedHash_ShouldReturnFalseAndZero()
        Dim Iterations As Integer = -1
        Dim Result As Boolean = PasswordHasher.TryGetIterations("invalid-hash", Iterations)
        Assert.IsFalse(Result)
        Assert.AreEqual(0, Iterations)
    End Sub
    <TestMethod>
    Public Sub HashPassword_NullPassword_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() PasswordHasher.HashPassword(Nothing, TestIterations))
    End Sub
    <TestMethod>
    Public Sub HashPassword_EmptyPassword_ShouldThrowArgumentException()
        Assert.ThrowsException(Of ArgumentException)(Sub() PasswordHasher.HashPassword(String.Empty, TestIterations))
    End Sub
    <TestMethod>
    Public Sub HashPassword_IterationsBelowMinimum_ShouldThrowArgumentOutOfRangeException()
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() PasswordHasher.HashPassword(TestPassword, PasswordHasher.MinimumIterations - 1))
    End Sub
    <TestMethod>
    Public Sub HashPassword_IterationsAboveMaximum_ShouldThrowArgumentOutOfRangeException()
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() PasswordHasher.HashPassword(TestPassword, PasswordHasher.MaximumIterations + 1))
    End Sub
    <TestMethod>
    Public Sub VerifyPasswordDetailed_InvalidRequiredIterations_ShouldThrowArgumentOutOfRangeException()
        Dim EncodedHash As String = PasswordHasher.HashPassword(TestPassword, TestIterations)
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() PasswordHasher.VerifyPasswordDetailed(TestPassword, EncodedHash, PasswordHasher.MinimumIterations - 1))
    End Sub
End Class