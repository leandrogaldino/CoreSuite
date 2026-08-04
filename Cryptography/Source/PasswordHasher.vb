Imports System.Globalization
Imports System.Security.Cryptography
''' <summary>
''' Provides versioned password hashing and verification using PBKDF2-HMAC-SHA256.
''' </summary>
Public NotInheritable Class PasswordHasher
    Public Const CurrentFormatVersion As Integer = 1
    Public Const DefaultIterations As Integer = 600_000
    Public Const MinimumIterations As Integer = 100_000
    Public Const MaximumIterations As Integer = 5_000_000
    Private Const FormatIdentifier As String = "CSPH"
    Private Const AlgorithmIdentifier As String = "PBKDF2-SHA256"
    Private Const SaltSize As Integer = 16
    Private Const HashSize As Integer = 32
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Creates a self-contained, versioned password hash containing the algorithm, iterations, salt, and derived hash.
    ''' </summary>
    ''' <param name="Password">The non-empty password to hash.</param>
    ''' <param name="Iterations">The PBKDF2 iteration count.</param>
    ''' <returns>A string that can be stored directly in a database.</returns>
    Public Shared Function HashPassword(Password As String, Optional Iterations As Integer = DefaultIterations) As String
        ValidatePassword(Password)
        ValidateIterations(Iterations)
        Dim Salt As Byte() = RandomNumberGenerator.GetBytes(SaltSize)
        Dim HashBytes As Byte() = DeriveHash(Password, Salt, Iterations)
        Try
            Return String.Join("$", FormatIdentifier, CurrentFormatVersion.ToString(CultureInfo.InvariantCulture), AlgorithmIdentifier, Iterations.ToString(CultureInfo.InvariantCulture), Convert.ToBase64String(Salt), Convert.ToBase64String(HashBytes))
        Finally
            CryptographicOperations.ZeroMemory(HashBytes.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Verifies a password against a stored versioned password hash.
    ''' </summary>
    ''' <param name="Password">The password to verify.</param>
    ''' <param name="EncodedHash">The stored hash produced by <see cref="HashPassword"/>.</param>
    ''' <returns><see langword="True"/> when the password matches; otherwise, <see langword="False"/>.</returns>
    Public Shared Function VerifyPassword(Password As String, EncodedHash As String) As Boolean
        Return VerifyPasswordDetailed(Password, EncodedHash) <> PasswordVerificationResult.Failed
    End Function
    ''' <summary>
    ''' Verifies a password and indicates whether the stored hash should be upgraded.
    ''' </summary>
    ''' <param name="Password">The password to verify.</param>
    ''' <param name="EncodedHash">The stored hash produced by <see cref="HashPassword"/>.</param>
    ''' <param name="RequiredIterations">The minimum desired PBKDF2 iteration count.</param>
    ''' <returns>A detailed password verification result.</returns>
    Public Shared Function VerifyPasswordDetailed(Password As String, EncodedHash As String, Optional RequiredIterations As Integer = DefaultIterations) As PasswordVerificationResult
        ValidatePassword(Password)
        ValidateIterations(RequiredIterations)
        Dim Iterations As Integer
        Dim Salt As Byte() = Nothing
        Dim StoredHash As Byte() = Nothing
        If Not TryParse(EncodedHash, Iterations, Salt, StoredHash) Then Return PasswordVerificationResult.Failed
        Dim ComputedHash As Byte() = DeriveHash(Password, Salt, Iterations)
        Try
            If Not CryptographicOperations.FixedTimeEquals(ComputedHash, StoredHash) Then Return PasswordVerificationResult.Failed
            If Iterations < RequiredIterations Then Return PasswordVerificationResult.SuccessRehashNeeded
            Return PasswordVerificationResult.Success
        Finally
            CryptographicOperations.ZeroMemory(ComputedHash.AsSpan())
            CryptographicOperations.ZeroMemory(StoredHash.AsSpan())
        End Try
    End Function
    ''' <summary>
    ''' Determines whether a stored password hash is malformed, unsupported, or uses fewer iterations than required.
    ''' </summary>
    ''' <param name="EncodedHash">The stored password hash.</param>
    ''' <param name="RequiredIterations">The minimum desired PBKDF2 iteration count.</param>
    ''' <returns><see langword="True"/> when the password should be hashed again after successful authentication.</returns>
    Public Shared Function NeedsRehash(EncodedHash As String, Optional RequiredIterations As Integer = DefaultIterations) As Boolean
        ValidateIterations(RequiredIterations)
        Dim Iterations As Integer
        Dim Salt As Byte() = Nothing
        Dim StoredHash As Byte() = Nothing
        If Not TryParse(EncodedHash, Iterations, Salt, StoredHash) Then Return True
        CryptographicOperations.ZeroMemory(StoredHash.AsSpan())
        Return Iterations < RequiredIterations
    End Function
    ''' <summary>
    ''' Attempts to obtain the PBKDF2 iteration count stored in a password hash.
    ''' </summary>
    ''' <param name="EncodedHash">The stored password hash.</param>
    ''' <param name="Iterations">Receives the stored iteration count.</param>
    ''' <returns><see langword="True"/> when the hash is valid and supported; otherwise, <see langword="False"/>.</returns>
    Public Shared Function TryGetIterations(EncodedHash As String, ByRef Iterations As Integer) As Boolean
        Dim Salt As Byte() = Nothing
        Dim StoredHash As Byte() = Nothing
        Dim Success As Boolean = TryParse(EncodedHash, Iterations, Salt, StoredHash)
        If StoredHash IsNot Nothing Then CryptographicOperations.ZeroMemory(StoredHash.AsSpan())
        Return Success
    End Function
    Private Shared Function TryParse(EncodedHash As String, ByRef Iterations As Integer, ByRef Salt As Byte(), ByRef StoredHash As Byte()) As Boolean
        Iterations = 0
        Salt = Nothing
        StoredHash = Nothing
        If String.IsNullOrWhiteSpace(EncodedHash) Then Return False
        Dim Parts As String() = EncodedHash.Split(New Char() {"$"c}, StringSplitOptions.None)
        If Parts.Length <> 6 Then Return False
        If Not String.Equals(Parts(0), FormatIdentifier, StringComparison.Ordinal) Then Return False
        If Not String.Equals(Parts(1), CurrentFormatVersion.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal) Then Return False
        If Not String.Equals(Parts(2), AlgorithmIdentifier, StringComparison.Ordinal) Then Return False
        If Not Integer.TryParse(Parts(3), NumberStyles.None, CultureInfo.InvariantCulture, Iterations) Then Return False
        If Iterations < MinimumIterations OrElse Iterations > MaximumIterations Then Return False
        Try
            Salt = Convert.FromBase64String(Parts(4))
            StoredHash = Convert.FromBase64String(Parts(5))
        Catch Ex As FormatException
            Salt = Nothing
            StoredHash = Nothing
            Return False
        End Try
        If Salt.Length <> SaltSize OrElse StoredHash.Length <> HashSize Then
            CryptographicOperations.ZeroMemory(StoredHash.AsSpan())
            Salt = Nothing
            StoredHash = Nothing
            Return False
        End If
        Return True
    End Function
    Private Shared Function DeriveHash(Password As String, Salt As Byte(), Iterations As Integer) As Byte()
        Return Rfc2898DeriveBytes.Pbkdf2(Password, Salt, Iterations, HashAlgorithmName.SHA256, HashSize)
    End Function
    Private Shared Sub ValidatePassword(Password As String)
        ArgumentNullException.ThrowIfNull(Password)
        If Password.Length = 0 Then Throw New ArgumentException("Password cannot be empty.", NameOf(Password))
    End Sub
    Private Shared Sub ValidateIterations(Iterations As Integer)
        If Iterations < MinimumIterations OrElse Iterations > MaximumIterations Then Throw New ArgumentOutOfRangeException(NameOf(Iterations), $"Iterations must be between {MinimumIterations:N0} and {MaximumIterations:N0}.")
    End Sub
End Class
