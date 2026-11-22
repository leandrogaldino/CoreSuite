''' <summary>
''' Groups authenticated encryption, SHA-256 hashing, and password hashing
''' under a single public entry point for .NET Framework 4.8.
''' </summary>
''' <remarks>
''' Use <c>Cryptography.TextEncryption.Encrypt(...)</c>,
''' <c>Cryptography.Hashing.ComputeSha256(...)</c>, or
''' <c>Cryptography.PasswordHasher.HashPassword(...)</c>.
''' The nested categories delegate to the existing implementations.
''' </remarks>
Public NotInheritable Partial Class Cryptography
    Private Sub New()
    End Sub
End Class
