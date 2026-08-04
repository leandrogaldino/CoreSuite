Imports System.Security.Cryptography
Imports System.Text
Imports CoreSuite.Services.Cryptography
Imports Microsoft.VisualStudio.TestTools.UnitTesting
<TestClass>
Public NotInheritable Class TextEncryptionTests
    Private Const TestIterations As Integer = TextEncryption.MinimumIterations
    Private Const TestPassword As String = "CoreSuite-Encryption-Key-2026!"
    <TestMethod>
    Public Sub EncryptAndDecrypt_UnicodeText_ShouldPreserveContent()
        Dim OriginalText As String = "CoreSuite — criptografia segura — ção — 🔐"
        Dim EncryptedText As String = TextEncryption.Encrypt(OriginalText, TestPassword, TestIterations)
        Dim DecryptedText As String = TextEncryption.Decrypt(EncryptedText, TestPassword)
        Assert.AreEqual(OriginalText, DecryptedText)
    End Sub
    <TestMethod>
    Public Sub EncryptAndDecrypt_EmptyText_ShouldPreserveEmptyContent()
        Dim EncryptedText As String = TextEncryption.Encrypt(String.Empty, TestPassword, TestIterations)
        Dim DecryptedText As String = TextEncryption.Decrypt(EncryptedText, TestPassword)
        Assert.AreEqual(String.Empty, DecryptedText)
    End Sub
    <TestMethod>
    Public Sub Encrypt_SameTextTwice_ShouldGenerateDifferentPackages()
        Dim FirstEncryptedText As String = TextEncryption.Encrypt("Same content", TestPassword, TestIterations)
        Dim SecondEncryptedText As String = TextEncryption.Encrypt("Same content", TestPassword, TestIterations)
        Assert.AreNotEqual(FirstEncryptedText, SecondEncryptedText)
        Assert.AreEqual("Same content", TextEncryption.Decrypt(FirstEncryptedText, TestPassword))
        Assert.AreEqual("Same content", TextEncryption.Decrypt(SecondEncryptedText, TestPassword))
    End Sub
    <TestMethod>
    Public Sub EncryptBytesAndDecryptBytes_BinaryData_ShouldPreserveContent()
        Dim OriginalData As Byte() = {&H0, &H1, &H2, &H7F, &H80, &HFE, &HFF}
        Dim EncryptedData As Byte() = TextEncryption.EncryptBytes(OriginalData, TestPassword, TestIterations)
        Dim DecryptedData As Byte() = TextEncryption.DecryptBytes(EncryptedData, TestPassword)
        CollectionAssert.AreEqual(OriginalData, DecryptedData)
    End Sub
    <TestMethod>
    Public Sub Decrypt_WrongPassword_ShouldThrowAuthenticationTagMismatchException()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Assert.ThrowsException(Of AuthenticationTagMismatchException)(Sub() TextEncryption.Decrypt(EncryptedText, "Wrong password"))
    End Sub
    <TestMethod>
    Public Sub Decrypt_TamperedPayload_ShouldThrowAuthenticationTagMismatchException()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim EncryptedData As Byte() = Convert.FromBase64String(EncryptedText)
        EncryptedData(EncryptedData.Length - 1) = EncryptedData(EncryptedData.Length - 1) Xor CByte(1)
        Dim TamperedText As String = Convert.ToBase64String(EncryptedData)
        Assert.ThrowsException(Of AuthenticationTagMismatchException)(Sub() TextEncryption.Decrypt(TamperedText, TestPassword))
    End Sub
    <TestMethod>
    Public Sub TryDecrypt_ValidPackage_ShouldReturnTrueAndContent()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim DecryptedText As String = Nothing
        Dim Result As Boolean = TextEncryption.TryDecrypt(EncryptedText, TestPassword, DecryptedText)
        Assert.IsTrue(Result)
        Assert.AreEqual("Protected content", DecryptedText)
    End Sub
    <TestMethod>
    Public Sub TryDecrypt_WrongPassword_ShouldReturnFalseAndEmptyText()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim DecryptedText As String = "Initial value"
        Dim Result As Boolean = TextEncryption.TryDecrypt(EncryptedText, "Wrong password", DecryptedText)
        Assert.IsFalse(Result)
        Assert.AreEqual(String.Empty, DecryptedText)
    End Sub
    <TestMethod>
    Public Sub TryDecrypt_InvalidUtf8Payload_ShouldReturnFalse()
        Dim InvalidUtf8 As Byte() = {&HFF}
        Dim EncryptedData As Byte() = TextEncryption.EncryptBytes(InvalidUtf8, TestPassword, TestIterations)
        Dim EncryptedText As String = Convert.ToBase64String(EncryptedData)
        Dim DecryptedText As String = Nothing
        Dim Result As Boolean = TextEncryption.TryDecrypt(EncryptedText, TestPassword, DecryptedText)
        Assert.IsFalse(Result)
        Assert.AreEqual(String.Empty, DecryptedText)
    End Sub
    <TestMethod>
    Public Sub TryDecryptBytes_InvalidPackage_ShouldReturnFalseAndEmptyData()
        Dim DecryptedData As Byte() = Nothing
        Dim Result As Boolean = TextEncryption.TryDecryptBytes(New Byte() {&H1, &H2, &H3}, TestPassword, DecryptedData)
        Assert.IsFalse(Result)
        Assert.IsNotNull(DecryptedData)
        Assert.AreEqual(0, DecryptedData.Length)
    End Sub
    <TestMethod>
    Public Sub IsEncrypted_ValidTextAndBinaryPackages_ShouldReturnTrue()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim EncryptedData As Byte() = Convert.FromBase64String(EncryptedText)
        Assert.IsTrue(TextEncryption.IsEncrypted(EncryptedText))
        Assert.IsTrue(TextEncryption.IsEncrypted(EncryptedData))
    End Sub
    <TestMethod>
    Public Sub IsEncrypted_InvalidValues_ShouldReturnFalse()
        Assert.IsFalse(TextEncryption.IsEncrypted(String.Empty))
        Assert.IsFalse(TextEncryption.IsEncrypted("not-base64"))
        Assert.IsFalse(TextEncryption.IsEncrypted(DirectCast(Nothing, Byte())))
        Assert.IsFalse(TextEncryption.IsEncrypted(New Byte() {&H1, &H2, &H3}))
    End Sub
    <TestMethod>
    Public Sub IsEncrypted_TamperedPayload_ShouldRemainStructurallyValid()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim EncryptedData As Byte() = Convert.FromBase64String(EncryptedText)
        EncryptedData(EncryptedData.Length - 1) = EncryptedData(EncryptedData.Length - 1) Xor CByte(1)
        Assert.IsTrue(TextEncryption.IsEncrypted(EncryptedData))
    End Sub
    <TestMethod>
    Public Sub GetIterations_ValidPackage_ShouldReturnStoredIterations()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Assert.AreEqual(TestIterations, TextEncryption.GetIterations(EncryptedText))
    End Sub
    <TestMethod>
    Public Sub NeedsReEncryption_ShouldCompareStoredAndRequiredIterations()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Assert.IsFalse(TextEncryption.NeedsReEncryption(EncryptedText, TestIterations))
        Assert.IsTrue(TextEncryption.NeedsReEncryption(EncryptedText, TestIterations + 1))
    End Sub
    <TestMethod>
    Public Sub Encrypt_NullText_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() TextEncryption.Encrypt(Nothing, TestPassword, TestIterations))
    End Sub
    <TestMethod>
    Public Sub Encrypt_EmptyPassword_ShouldThrowArgumentException()
        Assert.ThrowsException(Of ArgumentException)(Sub() TextEncryption.Encrypt("Content", String.Empty, TestIterations))
    End Sub
    <TestMethod>
    Public Sub Encrypt_NullPassword_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() TextEncryption.Encrypt("Content", Nothing, TestIterations))
    End Sub
    <TestMethod>
    Public Sub Encrypt_IterationsOutsideSupportedRange_ShouldThrowArgumentOutOfRangeException()
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() TextEncryption.Encrypt("Content", TestPassword, TextEncryption.MinimumIterations - 1))
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() TextEncryption.Encrypt("Content", TestPassword, TextEncryption.MaximumIterations + 1))
    End Sub
    <TestMethod>
    Public Sub Decrypt_InvalidBase64_ShouldThrowFormatException()
        Assert.ThrowsException(Of FormatException)(Sub() TextEncryption.Decrypt("not-base64", TestPassword))
    End Sub
    <TestMethod>
    Public Sub Decrypt_EmptyEncryptedText_ShouldThrowArgumentException()
        Assert.ThrowsException(Of ArgumentException)(Sub() TextEncryption.Decrypt(String.Empty, TestPassword))
    End Sub
    <TestMethod>
    Public Sub Decrypt_UnsupportedVersion_ShouldThrowFormatException()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim EncryptedData As Byte() = Convert.FromBase64String(EncryptedText)
        EncryptedData(4) = TextEncryption.CurrentFormatVersion + CByte(1)
        Dim UnsupportedPackage As String = Convert.ToBase64String(EncryptedData)
        Assert.ThrowsException(Of FormatException)(Sub() TextEncryption.Decrypt(UnsupportedPackage, TestPassword))
    End Sub
    <TestMethod>
    Public Sub Decrypt_InvalidPackageLength_ShouldThrowFormatException()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Dim EncryptedData As Byte() = Convert.FromBase64String(EncryptedText)
        Array.Resize(EncryptedData, EncryptedData.Length + 1)
        Dim InvalidPackage As String = Convert.ToBase64String(EncryptedData)
        Assert.ThrowsException(Of FormatException)(Sub() TextEncryption.Decrypt(InvalidPackage, TestPassword))
    End Sub
    <TestMethod>
    Public Sub DecryptBytes_NullPackage_ShouldThrowArgumentNullException()
        Assert.ThrowsException(Of ArgumentNullException)(Sub() TextEncryption.DecryptBytes(Nothing, TestPassword))
    End Sub
    <TestMethod>
    Public Sub NeedsReEncryption_InvalidRequiredIterations_ShouldThrowArgumentOutOfRangeException()
        Dim EncryptedText As String = TextEncryption.Encrypt("Protected content", TestPassword, TestIterations)
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() TextEncryption.NeedsReEncryption(EncryptedText, TextEncryption.MinimumIterations - 1))
    End Sub
End Class