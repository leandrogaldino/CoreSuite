Imports System.Runtime.CompilerServices
Imports System.Security.Cryptography
Imports System.Text

''' <summary>Provides shared cryptographic buffer, randomness and key derivation operations.</summary>
Public NotInheritable Class Helper
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

    ''' <summary>Derives a key using PBKDF2 and strict UTF-8 password encoding.</summary>
    Public Shared Function Pbkdf2(Password As String, Salt As Byte(), Iterations As Integer, Algorithm As HashAlgorithmName, Length As Integer) As Byte()
        ' Reject malformed UTF-16 instead of silently changing the password.
        Dim PasswordBytes = New UTF8Encoding(False, True).GetBytes(Password)
        Try
            Using Deriver As New Rfc2898DeriveBytes(PasswordBytes, Salt, Iterations, Algorithm)
                Return Deriver.GetBytes(Length)
            End Using
        Finally
            ZeroMemory(PasswordBytes)
        End Try
    End Function
End Class
