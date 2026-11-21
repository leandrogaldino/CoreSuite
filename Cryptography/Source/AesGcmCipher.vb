Imports CoreSuite.Helpers
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Crypto.Engines
Imports Org.BouncyCastle.Crypto.Modes
Imports Org.BouncyCastle.Crypto.Parameters
Imports System.Security.Cryptography

Friend NotInheritable Class AesGcmCipher
    Implements IDisposable
    Private ReadOnly _Key As Byte()
    Private ReadOnly _TagSize As Integer
    Public Sub New(Key As Byte(), TagSize As Integer)
        _Key = DirectCast(Key.Clone(), Byte())
        _TagSize = TagSize
    End Sub
    Public Sub Encrypt(Nonce As Byte(), Plaintext As Byte(), Ciphertext As Byte(), Tag As Byte(), Header As Byte())
        Dim Cipher = New GcmBlockCipher(New AesEngine())
        Cipher.Init(True, New AeadParameters(New KeyParameter(_Key), _TagSize * 8, Nonce, Header))
        Dim Output(Cipher.GetOutputSize(Plaintext.Length) - 1) As Byte
        Try
            Dim Count = Cipher.ProcessBytes(Plaintext, 0, Plaintext.Length, Output, 0)
            Cipher.DoFinal(Output, Count)
            Buffer.BlockCopy(Output, 0, Ciphertext, 0, Ciphertext.Length)
            Buffer.BlockCopy(Output, Ciphertext.Length, Tag, 0, Tag.Length)
        Finally
            CryptographyHelper.ZeroMemory(Output)
        End Try
    End Sub
    Public Sub Decrypt(Nonce As Byte(), Ciphertext As Byte(), Tag As Byte(), Plaintext As Byte(), Header As Byte())
        Dim Cipher = New GcmBlockCipher(New AesEngine())
        Cipher.Init(False, New AeadParameters(New KeyParameter(_Key), _TagSize * 8, Nonce, Header))
        Dim Input(Ciphertext.Length + Tag.Length - 1) As Byte
        Buffer.BlockCopy(Ciphertext, 0, Input, 0, Ciphertext.Length)
        Buffer.BlockCopy(Tag, 0, Input, Ciphertext.Length, Tag.Length)
        Dim Output(Cipher.GetOutputSize(Input.Length) - 1) As Byte
        Try
            Dim Count = Cipher.ProcessBytes(Input, 0, Input.Length, Output, 0)
            Cipher.DoFinal(Output, Count)
            Buffer.BlockCopy(Output, 0, Plaintext, 0, Plaintext.Length)
        Catch Ex As InvalidCipherTextException
            CryptographyHelper.ZeroMemory(Plaintext)
            Throw New AuthenticationTagMismatchException("The authentication tag did not match.", Ex)
        Finally
            CryptographyHelper.ZeroMemory(Output)
            CryptographyHelper.ZeroMemory(Input)
        End Try
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        CryptographyHelper.ZeroMemory(_Key)
    End Sub
End Class
