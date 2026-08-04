Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
''' <summary>
''' Provides password-based authenticated encryption for text and binary data using AES-256-GCM and PBKDF2-HMAC-SHA256.
''' </summary>
Public NotInheritable Class TextEncryption
    Public Const CurrentFormatVersion As Byte = 1
    Public Const DefaultIterations As Integer = 600_000
    Public Const MinimumIterations As Integer = 100_000
    Public Const MaximumIterations As Integer = 5_000_000
    Private Const MagicSize As Integer = 4
    Private Const HeaderSize As Integer = 18
    Private Const SaltSize As Integer = 16
    Private Const NonceSize As Integer = 12
    Private Const TagSize As Integer = 16
    Private Const KeySize As Integer = 32
    Private Const KdfIdentifier As Byte = 1
    Private Const CipherIdentifier As Byte = 1
    Private Shared ReadOnly MagicBytes As Byte() = {&H43, &H53, &H45, &H43}
    Private Shared ReadOnly Utf8 As New UTF8Encoding(False, True)
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Encrypts text using a password and returns the complete encrypted package encoded as Base64.
    ''' </summary>
    ''' <param name="Text">The text to encrypt. Empty text is supported.</param>
    ''' <param name="Password">The non-empty password used to derive the encryption key.</param>
    ''' <param name="Iterations">The PBKDF2 iteration count stored in the encrypted package.</param>
    ''' <returns>A versioned Base64 string containing the salt, nonce, authentication tag, metadata, and encrypted content.</returns>
    Public Shared Function Encrypt(Text As String, Password As String, Optional Iterations As Integer = DefaultIterations) As String
        ArgumentNullException.ThrowIfNull(Text)
        Dim DataBytes As Byte() = Utf8.GetBytes(Text)
        Try
            Return Convert.ToBase64String(EncryptBytes(DataBytes, Password, Iterations))
        Finally
            CryptographicOperations.ZeroMemory(DataBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Decrypts a Base64 encrypted package and returns its UTF-8 text.
    ''' </summary>
    ''' <param name="EncryptedText">The encrypted package produced by <see cref="Encrypt"/>.</param>
    ''' <param name="Password">The password used when the package was encrypted.</param>
    ''' <returns>The decrypted text.</returns>
    ''' <exception cref="FormatException">The encrypted package is malformed or uses an unsupported format.</exception>
    ''' <exception cref="CryptographicException">The password is incorrect or the encrypted package was modified.</exception>
    Public Shared Function Decrypt(EncryptedText As String, Password As String) As String
        If String.IsNullOrWhiteSpace(EncryptedText) Then Throw New ArgumentException("Encrypted text cannot be empty.", NameOf(EncryptedText))
        Dim EncryptedBytes As Byte()
        Try
            EncryptedBytes = Convert.FromBase64String(EncryptedText)
        Catch Ex As FormatException
            Throw New FormatException("The encrypted text is not valid Base64.", Ex)
        End Try
        Dim DataBytes As Byte() = DecryptBytes(EncryptedBytes, Password)
        Try
            Return Utf8.GetString(DataBytes)
        Finally
            CryptographicOperations.ZeroMemory(DataBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Attempts to decrypt a Base64 encrypted package without throwing for malformed data, authentication failure, or an incorrect password.
    ''' </summary>
    ''' <param name="EncryptedText">The encrypted package to decrypt.</param>
    ''' <param name="Password">The password used when the package was encrypted.</param>
    ''' <param name="Text">Receives the decrypted text when the operation succeeds.</param>
    ''' <returns><see langword="True"/> when decryption succeeds; otherwise, <see langword="False"/>.</returns>
    Public Shared Function TryDecrypt(EncryptedText As String, Password As String, ByRef Text As String) As Boolean
        Text = String.Empty
        If String.IsNullOrWhiteSpace(EncryptedText) OrElse String.IsNullOrEmpty(Password) Then Return False
        Try
            Text = Decrypt(EncryptedText, Password)
            Return True
        Catch Ex As FormatException
            Return False
        Catch Ex As CryptographicException
            Return False
        Catch Ex As ArgumentException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Encrypts binary data using a password and returns a versioned authenticated binary package.
    ''' </summary>
    ''' <param name="Data">The binary data to encrypt.</param>
    ''' <param name="Password">The non-empty password used to derive the encryption key.</param>
    ''' <param name="Iterations">The PBKDF2 iteration count stored in the encrypted package.</param>
    ''' <returns>The complete encrypted binary package.</returns>
    Public Shared Function EncryptBytes(Data As Byte(), Password As String, Optional Iterations As Integer = DefaultIterations) As Byte()
        ArgumentNullException.ThrowIfNull(Data)
        ValidatePassword(Password)
        ValidateIterations(Iterations)
        Dim Salt As Byte() = RandomNumberGenerator.GetBytes(SaltSize)
        Dim Nonce As Byte() = RandomNumberGenerator.GetBytes(NonceSize)
        Dim Tag(TagSize - 1) As Byte
        Dim CipherText As Byte() = CreateBuffer(Data.Length)
        Dim Header As Byte() = CreateHeader(Iterations, Data.Length)
        Dim KeyBytes As Byte() = DeriveKey(Password, Salt, Iterations)
        Try
            Using Algorithm As New AesGcm(KeyBytes, TagSize)
                Algorithm.Encrypt(Nonce, Data, CipherText, Tag, Header)
            End Using
            Dim Package(Header.Length + Salt.Length + Nonce.Length + Tag.Length + CipherText.Length - 1) As Byte
            Dim Offset As Integer = 0
            CopyTo(Header, Package, Offset)
            CopyTo(Salt, Package, Offset)
            CopyTo(Nonce, Package, Offset)
            CopyTo(Tag, Package, Offset)
            CopyTo(CipherText, Package, Offset)
            Return Package
        Finally
            CryptographicOperations.ZeroMemory(KeyBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Decrypts a versioned authenticated binary package using the supplied password.
    ''' </summary>
    ''' <param name="EncryptedData">The encrypted package produced by <see cref="EncryptBytes"/>.</param>
    ''' <param name="Password">The password used when the package was encrypted.</param>
    ''' <returns>The decrypted binary data.</returns>
    ''' <exception cref="FormatException">The package is malformed or uses an unsupported format.</exception>
    ''' <exception cref="CryptographicException">The password is incorrect or the package was modified.</exception>
    Public Shared Function DecryptBytes(EncryptedData As Byte(), Password As String) As Byte()
        ArgumentNullException.ThrowIfNull(EncryptedData)
        ValidatePassword(Password)
        Dim Header As Byte() = Nothing
        Dim Iterations As Integer
        Dim Salt As Byte() = Nothing
        Dim Nonce As Byte() = Nothing
        Dim Tag As Byte() = Nothing
        Dim CipherText As Byte() = Nothing
        ParsePackage(EncryptedData, Header, Iterations, Salt, Nonce, Tag, CipherText)
        Dim DataBytes As Byte() = CreateBuffer(CipherText.Length)
        Dim KeyBytes As Byte() = DeriveKey(Password, Salt, Iterations)
        Try
            Using Algorithm As New AesGcm(KeyBytes, TagSize)
                Algorithm.Decrypt(Nonce, CipherText, Tag, DataBytes, Header)
            End Using
            Return DataBytes
        Catch
            CryptographicOperations.ZeroMemory(DataBytes.AsSpan())
            Throw
        Finally
            CryptographicOperations.ZeroMemory(KeyBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Attempts to decrypt an authenticated binary package without throwing for malformed data, authentication failure, or an incorrect password.
    ''' </summary>
    ''' <param name="EncryptedData">The encrypted binary package.</param>
    ''' <param name="Password">The password used when the package was encrypted.</param>
    ''' <param name="Data">Receives the decrypted bytes when the operation succeeds.</param>
    ''' <returns><see langword="True"/> when decryption succeeds; otherwise, <see langword="False"/>.</returns>
    Public Shared Function TryDecryptBytes(EncryptedData As Byte(), Password As String, ByRef Data As Byte()) As Boolean
        Data = Array.Empty(Of Byte)()
        If EncryptedData Is Nothing OrElse String.IsNullOrEmpty(Password) Then Return False
        Try
            Data = DecryptBytes(EncryptedData, Password)
            Return True
        Catch Ex As FormatException
            Return False
        Catch Ex As CryptographicException
            Return False
        Catch Ex As ArgumentException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Determines whether a string contains a structurally valid encrypted package created by this class.
    ''' </summary>
    ''' <param name="EncryptedText">The Base64 text to inspect.</param>
    ''' <returns><see langword="True"/> when the package structure is valid; otherwise, <see langword="False"/>.</returns>
    Public Shared Function IsEncrypted(EncryptedText As String) As Boolean
        If String.IsNullOrWhiteSpace(EncryptedText) Then Return False
        Try
            Return IsEncrypted(Convert.FromBase64String(EncryptedText))
        Catch Ex As FormatException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Determines whether binary data contains a structurally valid encrypted package created by this class.
    ''' </summary>
    ''' <param name="EncryptedData">The binary package to inspect.</param>
    ''' <returns><see langword="True"/> when the package structure is valid; otherwise, <see langword="False"/>.</returns>
    Public Shared Function IsEncrypted(EncryptedData As Byte()) As Boolean
        If EncryptedData Is Nothing Then Return False
        Try
            Dim Header As Byte() = Nothing
            Dim Iterations As Integer
            Dim Salt As Byte() = Nothing
            Dim Nonce As Byte() = Nothing
            Dim Tag As Byte() = Nothing
            Dim CipherText As Byte() = Nothing
            ParsePackage(EncryptedData, Header, Iterations, Salt, Nonce, Tag, CipherText)
            Return True
        Catch Ex As FormatException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Gets the PBKDF2 iteration count stored in an encrypted text package.
    ''' </summary>
    ''' <param name="EncryptedText">The encrypted text package.</param>
    ''' <returns>The stored PBKDF2 iteration count.</returns>
    Public Shared Function GetIterations(EncryptedText As String) As Integer
        If String.IsNullOrWhiteSpace(EncryptedText) Then Throw New ArgumentException("Encrypted text cannot be empty.", NameOf(EncryptedText))
        Dim EncryptedData As Byte()
        Try
            EncryptedData = Convert.FromBase64String(EncryptedText)
        Catch Ex As FormatException
            Throw New FormatException("The encrypted text is not valid Base64.", Ex)
        End Try
        Dim Header As Byte() = Nothing
        Dim Iterations As Integer
        Dim Salt As Byte() = Nothing
        Dim Nonce As Byte() = Nothing
        Dim Tag As Byte() = Nothing
        Dim CipherText As Byte() = Nothing
        ParsePackage(EncryptedData, Header, Iterations, Salt, Nonce, Tag, CipherText)
        Return Iterations
    End Function
    ''' <summary>
    ''' Determines whether an encrypted text package should be encrypted again using a higher PBKDF2 iteration count.
    ''' </summary>
    ''' <param name="EncryptedText">The encrypted text package.</param>
    ''' <param name="RequiredIterations">The minimum desired PBKDF2 iteration count.</param>
    ''' <returns><see langword="True"/> when the stored iteration count is lower than the required value.</returns>
    Public Shared Function NeedsReEncryption(EncryptedText As String, Optional RequiredIterations As Integer = DefaultIterations) As Boolean
        ValidateIterations(RequiredIterations)
        Return GetIterations(EncryptedText) < RequiredIterations
    End Function
    Private Shared Function CreateHeader(Iterations As Integer, CipherTextLength As Integer) As Byte()
        Using Stream As New MemoryStream(HeaderSize)
            Using Writer As New BinaryWriter(Stream, Encoding.UTF8, True)
                Writer.Write(MagicBytes)
                Writer.Write(CurrentFormatVersion)
                Writer.Write(KdfIdentifier)
                Writer.Write(CipherIdentifier)
                Writer.Write(Iterations)
                Writer.Write(CByte(SaltSize))
                Writer.Write(CByte(NonceSize))
                Writer.Write(CByte(TagSize))
                Writer.Write(CipherTextLength)
            End Using
            Dim Header As Byte() = Stream.ToArray()
            If Header.Length <> HeaderSize Then Throw New InvalidOperationException("The encryption header has an unexpected size.")
            Return Header
        End Using
    End Function
    Private Shared Sub ParsePackage(EncryptedData As Byte(), ByRef Header As Byte(), ByRef Iterations As Integer, ByRef Salt As Byte(), ByRef Nonce As Byte(), ByRef Tag As Byte(), ByRef CipherText As Byte())
        Dim MinimumLength As Integer = HeaderSize + SaltSize + NonceSize + TagSize
        If EncryptedData.Length < MinimumLength Then Throw New FormatException("The encrypted package is incomplete.")
        Dim Version As Byte
        Dim Kdf As Byte
        Dim Cipher As Byte
        Dim StoredSaltSize As Byte
        Dim StoredNonceSize As Byte
        Dim StoredTagSize As Byte
        Dim CipherTextLength As Integer
        Using Stream As New MemoryStream(EncryptedData, False)
            Using Reader As New BinaryReader(Stream, Encoding.UTF8, True)
                Dim StoredMagic As Byte() = Reader.ReadBytes(MagicSize)
                If Not ByteArraysEqual(StoredMagic, MagicBytes) Then Throw New FormatException("The encrypted package identifier is invalid.")
                Version = Reader.ReadByte()
                Kdf = Reader.ReadByte()
                Cipher = Reader.ReadByte()
                Iterations = Reader.ReadInt32()
                StoredSaltSize = Reader.ReadByte()
                StoredNonceSize = Reader.ReadByte()
                StoredTagSize = Reader.ReadByte()
                CipherTextLength = Reader.ReadInt32()
            End Using
        End Using
        If Version <> CurrentFormatVersion Then Throw New FormatException($"Encryption format version {Version} is not supported.")
        If Kdf <> KdfIdentifier Then Throw New FormatException("The key derivation algorithm is not supported.")
        If Cipher <> CipherIdentifier Then Throw New FormatException("The encryption algorithm is not supported.")
        ValidateIterationsForFormat(Iterations)
        If StoredSaltSize <> SaltSize OrElse StoredNonceSize <> NonceSize OrElse StoredTagSize <> TagSize Then Throw New FormatException("The encrypted package contains unsupported cryptographic parameter sizes.")
        If CipherTextLength < 0 Then Throw New FormatException("The encrypted package contains an invalid content length.")
        Dim ExpectedLength As Long = CLng(HeaderSize) + SaltSize + NonceSize + TagSize + CipherTextLength
        If ExpectedLength <> EncryptedData.LongLength Then Throw New FormatException("The encrypted package length is invalid.")
        Header = CopyBytes(EncryptedData, 0, HeaderSize)
        Dim Offset As Integer = HeaderSize
        Salt = CopyBytes(EncryptedData, Offset, SaltSize)
        Offset += SaltSize
        Nonce = CopyBytes(EncryptedData, Offset, NonceSize)
        Offset += NonceSize
        Tag = CopyBytes(EncryptedData, Offset, TagSize)
        Offset += TagSize
        CipherText = CopyBytes(EncryptedData, Offset, CipherTextLength)
    End Sub
    Private Shared Function DeriveKey(Password As String, Salt As Byte(), Iterations As Integer) As Byte()
        Return Rfc2898DeriveBytes.Pbkdf2(Password, Salt, Iterations, HashAlgorithmName.SHA256, KeySize)
    End Function
    Private Shared Sub ValidatePassword(Password As String)
        ArgumentNullException.ThrowIfNull(Password)
        If Password.Length = 0 Then Throw New ArgumentException("Password cannot be empty.", NameOf(Password))
    End Sub
    Private Shared Sub ValidateIterations(Iterations As Integer)
        If Iterations < MinimumIterations OrElse Iterations > MaximumIterations Then Throw New ArgumentOutOfRangeException(NameOf(Iterations), $"Iterations must be between {MinimumIterations:N0} and {MaximumIterations:N0}.")
    End Sub
    Private Shared Sub ValidateIterationsForFormat(Iterations As Integer)
        If Iterations < MinimumIterations OrElse Iterations > MaximumIterations Then Throw New FormatException("The encrypted package contains an unsupported PBKDF2 iteration count.")
    End Sub
    Private Shared Function CreateBuffer(Length As Integer) As Byte()
        If Length = 0 Then Return Array.Empty(Of Byte)()
        Return New Byte(Length - 1) {}
    End Function
    Private Shared Function CopyBytes(Source As Byte(), Offset As Integer, Count As Integer) As Byte()
        If Count = 0 Then Return Array.Empty(Of Byte)()
        Dim Result(Count - 1) As Byte
        Buffer.BlockCopy(Source, Offset, Result, 0, Count)
        Return Result
    End Function
    Private Shared Sub CopyTo(Source As Byte(), Destination As Byte(), ByRef Offset As Integer)
        If Source.Length > 0 Then Buffer.BlockCopy(Source, 0, Destination, Offset, Source.Length)
        Offset += Source.Length
    End Sub
    Private Shared Function ByteArraysEqual(First As Byte(), Second As Byte()) As Boolean
        If First.Length <> Second.Length Then Return False
        For Index As Integer = 0 To First.Length - 1
            If First(Index) <> Second(Index) Then Return False
        Next
        Return True
    End Function
End Class