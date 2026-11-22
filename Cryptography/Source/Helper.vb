Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text

''' <summary>Provides shared cryptographic buffer, randomness and key derivation operations.</summary>
Public NotInheritable Class Helper
    Private Const BCryptAlgHandleHmacFlag As Integer = &H8

    Private Sub New()
    End Sub

    ''' <summary>Clears a byte array without allowing the clearing call to be optimized away.</summary>
    <MethodImpl(MethodImplOptions.NoInlining Or MethodImplOptions.NoOptimization)>
    Public Shared Sub ZeroMemory(Bytes As Byte())
        If Bytes Is Nothing Then Throw New ArgumentNullException(NameOf(Bytes))
        Array.Clear(Bytes, 0, Bytes.Length)
    End Sub

    ''' <summary>Compares equal-length byte arrays without content-dependent early exit.</summary>
    <MethodImpl(MethodImplOptions.NoInlining Or MethodImplOptions.NoOptimization)>
    Public Shared Function FixedTimeEquals(Left As Byte(), Right As Byte()) As Boolean
        If Left Is Nothing Then Throw New ArgumentNullException(NameOf(Left))
        If Right Is Nothing Then Throw New ArgumentNullException(NameOf(Right))
        If Left.Length <> Right.Length Then Return False
        Dim Difference As Integer = 0
        For I = 0 To Left.Length - 1
            Difference = Difference Or (CInt(Left(I)) Xor CInt(Right(I)))
        Next
        Return Difference = 0
    End Function

    ''' <summary>Generates cryptographically strong random bytes.</summary>
    Public Shared Function RandomBytes(Count As Integer) As Byte()
        If Count < 0 Then Throw New ArgumentOutOfRangeException(NameOf(Count))
        Dim Bytes(Count - 1) As Byte
        Using Random = RandomNumberGenerator.Create()
            Random.GetBytes(Bytes)
        End Using
        Return Bytes
    End Function

    ''' <summary>
    ''' Derives PBKDF2 keys using the Windows CNG provider instead of the slower
    ''' Rfc2898DeriveBytes implementation in .NET Framework 4.8.
    ''' The result is byte-for-byte compatible for the same input parameters.
    ''' </summary>
    Public Shared Function Pbkdf2(Password As String, Salt As Byte(), Iterations As Integer, Algorithm As HashAlgorithmName, Length As Integer) As Byte()
        If Password Is Nothing Then Throw New ArgumentNullException(NameOf(Password))
        If Salt Is Nothing Then Throw New ArgumentNullException(NameOf(Salt))
        If Salt.Length < 8 Then Throw New ArgumentException("Salt must contain at least 8 bytes.", NameOf(Salt))
        If Iterations <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(Iterations))
        If Length <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(Length))

        Select Case Algorithm.Name
            Case HashAlgorithmName.SHA1.Name, HashAlgorithmName.SHA256.Name, HashAlgorithmName.SHA384.Name, HashAlgorithmName.SHA512.Name
                ' Algorithm identifiers are the same for HashAlgorithmName and Windows CNG.
            Case Else
                Throw New CryptographicException("Unsupported PBKDF2 hash algorithm.")
        End Select

        ' Preserve the previous strict UTF-8 conversion for encrypted-data compatibility.
        Dim PasswordBytes As Byte() = New UTF8Encoding(False, True).GetBytes(Password)
        Try
            Dim AlgorithmHandle As IntPtr = IntPtr.Zero
            Dim DerivedKey(Length - 1) As Byte
            Try
                Dim Status As Integer = BCryptOpenAlgorithmProvider(AlgorithmHandle, Algorithm.Name, IntPtr.Zero, BCryptAlgHandleHmacFlag)
                If Status <> 0 Then Throw New CryptographicException("BCryptOpenAlgorithmProvider failed (NTSTATUS 0x" & Status.ToString("X8") & ").")

                Status = BCryptDeriveKeyPBKDF2(AlgorithmHandle, PasswordBytes, PasswordBytes.Length, Salt, Salt.Length, CULng(Iterations), DerivedKey, DerivedKey.Length, 0)
                If Status <> 0 Then Throw New CryptographicException("BCryptDeriveKeyPBKDF2 failed (NTSTATUS 0x" & Status.ToString("X8") & ").")

                Return DerivedKey
            Catch
                ZeroMemory(DerivedKey)
                Throw
            Finally
                If AlgorithmHandle <> IntPtr.Zero Then BCryptCloseAlgorithmProvider(AlgorithmHandle, 0)
            End Try
        Finally
            ZeroMemory(PasswordBytes)
        End Try
    End Function

    <DllImport("bcrypt.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Private Shared Function BCryptOpenAlgorithmProvider(ByRef AlgorithmHandle As IntPtr, ByVal AlgorithmId As String, ByVal Implementation As IntPtr, ByVal Flags As Integer) As Integer
    End Function

    <DllImport("bcrypt.dll", ExactSpelling:=True)>
    Private Shared Function BCryptDeriveKeyPBKDF2(ByVal AlgorithmHandle As IntPtr, ByVal PasswordBytes As Byte(), ByVal PasswordLength As Integer, ByVal Salt As Byte(), ByVal SaltLength As Integer, ByVal Iterations As ULong, <Out()> ByVal DerivedKey As Byte(), ByVal DerivedKeyLength As Integer, ByVal Flags As Integer) As Integer
    End Function

    <DllImport("bcrypt.dll", ExactSpelling:=True)>
    Private Shared Function BCryptCloseAlgorithmProvider(ByVal AlgorithmHandle As IntPtr, ByVal Flags As Integer) As Integer
    End Function
End Class
