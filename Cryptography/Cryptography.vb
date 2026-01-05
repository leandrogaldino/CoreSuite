Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web.Script.Serialization
Public Class Cryptography
    Private Shared ReadOnly Salt() As Byte = {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H5, &H4, &H3, &H2, &H1, &H0}

    Public Shared Async Function EncryptAsync(Text As String, Key As String) As Task(Of String)
        If Not String.IsNullOrEmpty(Text) Then
            Dim CipherText As Byte()
            Dim EncryptKey As New Rfc2898DeriveBytes(Key, Salt)
            Dim Algorithm = New RijndaelManaged With {
                .Key = EncryptKey.GetBytes(16),
                .IV = EncryptKey.GetBytes(16)
            }
            Dim SourceBytes() As Byte = New UnicodeEncoding().GetBytes(Text)
            Using StreamSource = New MemoryStream(SourceBytes)
                Using StreamDestination As New MemoryStream()
                    Using Crypto As New CryptoStream(StreamSource, Algorithm.CreateEncryptor(), CryptoStreamMode.Read)
                        Await Crypto.CopyToAsync(StreamDestination)

                        CipherText = StreamDestination.ToArray()
                    End Using
                End Using
            End Using
            Return Convert.ToBase64String(CipherText)
        Else
            Return String.Empty
        End If
    End Function

    Public Shared Function Encrypt(Text As String, Key As String) As String
        If Not String.IsNullOrEmpty(Text) Then
            Dim CipherText As Byte()
            Dim EncryptKey As New Rfc2898DeriveBytes(Key, Salt)
            Dim Algorithm = New RijndaelManaged With {
                .Key = EncryptKey.GetBytes(16),
                .IV = EncryptKey.GetBytes(16)
            }
            Dim SourceBytes() As Byte = New UnicodeEncoding().GetBytes(Text)
            Using StreamSource = New MemoryStream(SourceBytes)
                Using StreamDestination As New MemoryStream()
                    Using Crypto As New CryptoStream(StreamSource, Algorithm.CreateEncryptor(), CryptoStreamMode.Read)
                        Crypto.CopyTo(StreamDestination)

                        CipherText = StreamDestination.ToArray()
                    End Using
                End Using
            End Using
            Return Convert.ToBase64String(CipherText)
        Else
            Return String.Empty
        End If
    End Function

    Public Shared Async Function DecryptAsync(Text As String, Key As String) As Task(Of String)
        If Not String.IsNullOrEmpty(Text) Then
            Dim DecryptKey As New Rfc2898DeriveBytes(Key, Salt)
            Dim Algorithm = New RijndaelManaged With {
                .Key = DecryptKey.GetBytes(16),
                .IV = DecryptKey.GetBytes(16)
            }
            Using StreamSource = New MemoryStream(Convert.FromBase64String(Text))
                Using StreamDestination As New MemoryStream()
                    Using Crypto As New CryptoStream(StreamSource, Algorithm.CreateDecryptor(), CryptoStreamMode.Read)
                        Await Crypto.CopyToAsync(StreamDestination)
                        Dim DecryptedBytes() As Byte = StreamDestination.ToArray()
                        Dim DecryptedText = New UnicodeEncoding().GetString(DecryptedBytes)
                        Return DecryptedText
                    End Using
                End Using
            End Using
        Else
            Return String.Empty
        End If
    End Function
    Public Shared Function Decrypt(Text As String, Key As String) As String
        If Not String.IsNullOrEmpty(Text) Then
            Dim DecryptKey As New Rfc2898DeriveBytes(Key, Salt)
            Dim Algorithm = New RijndaelManaged With {
                .Key = DecryptKey.GetBytes(16),
                .IV = DecryptKey.GetBytes(16)
            }
            Using StreamSource = New MemoryStream(Convert.FromBase64String(Text))
                Using StreamDestination As New MemoryStream()
                    Using Crypto As New CryptoStream(StreamSource, Algorithm.CreateDecryptor(), CryptoStreamMode.Read)
                        Crypto.CopyTo(StreamDestination)
                        Dim DecryptedBytes() As Byte = StreamDestination.ToArray()
                        Dim DecryptedText = New UnicodeEncoding().GetString(DecryptedBytes)
                        Return DecryptedText
                    End Using
                End Using
            End Using
        Else
            Return String.Empty
        End If

    End Function
End Class
Public Class Hashing
    Public Shared Function GetSHA256Hash(Text As String) As String
        Using Sha As SHA256 = SHA256.Create()
            Dim Bytes = Sha.ComputeHash(Encoding.UTF8.GetBytes(Text))
            Return String.Concat(Bytes.Select(Function(b) b.ToString("X2")))
        End Using
    End Function
    Public Shared Function HashPassword(Password As String, ByRef Salt As String, Optional Iterations As Integer = 100_000) As String
        Dim SaltBytes(15) As Byte
        Using Rng = RandomNumberGenerator.Create()
            Rng.GetBytes(SaltBytes)
        End Using
        Salt = Convert.ToBase64String(SaltBytes)
        Using Pbkdf2 As New Rfc2898DeriveBytes(Password, SaltBytes, Iterations, HashAlgorithmName.SHA256)
            Dim Hash = Pbkdf2.GetBytes(32)
            Return Convert.ToBase64String(Hash)
        End Using
    End Function
    Public Shared Function VerifyPassword(Password As String, Salt As String, StoredHash As String, Optional Iterations As Integer = 100_000) As Boolean
        Dim SaltBytes = Convert.FromBase64String(Salt)
        Using Pbkdf2 As New Rfc2898DeriveBytes(Password, SaltBytes, Iterations, HashAlgorithmName.SHA256)
            Dim Hash = Pbkdf2.GetBytes(32)
            Return Convert.ToBase64String(Hash) = StoredHash
        End Using
    End Function
End Class

Public Class SecureStorage
    Public Shared Sub Save(FilePath As String, Value As Dictionary(Of String, Object), Optional Scope As DataProtectionScope = DataProtectionScope.CurrentUser)
        If Value Is Nothing Then
            Throw New ArgumentNullException(NameOf(Value))
        End If
        Dim Folder = Path.GetDirectoryName(FilePath)
        If Not String.IsNullOrEmpty(Folder) AndAlso Not Directory.Exists(Folder) Then
            Directory.CreateDirectory(Folder)
        End If
        Dim Jss As New JavaScriptSerializer With {
            .MaxJsonLength = Integer.MaxValue
        }
        Dim Json As String = Jss.Serialize(Value)
        Dim DataBytes As Byte() = Encoding.UTF8.GetBytes(Json)
        Dim ProtectedDataBytes As Byte() =
        ProtectedData.Protect(DataBytes, Nothing, Scope)
        File.WriteAllBytes(FilePath, ProtectedDataBytes)
    End Sub
    Public Shared Function Load(FilePath As String, Optional Scope As DataProtectionScope = DataProtectionScope.CurrentUser) As Dictionary(Of String, Object)
        If Not File.Exists(FilePath) Then
            Return New Dictionary(Of String, Object)
        End If
        Try
            Dim ProtectedDataBytes As Byte() = File.ReadAllBytes(FilePath)
            Dim DataBytes As Byte() =
            ProtectedData.Unprotect(ProtectedDataBytes, Nothing, Scope)
            Dim Json As String = Encoding.UTF8.GetString(DataBytes)
            If String.IsNullOrWhiteSpace(Json) Then
                Return New Dictionary(Of String, Object)
            End If

            Dim Jss As New JavaScriptSerializer With {
                .MaxJsonLength = Integer.MaxValue
            }
            Dim Result = Jss.Deserialize(Of Dictionary(Of String, Object))(Json)
            Return If(Result, New Dictionary(Of String, Object))
        Catch ex As CryptographicException
            Return New Dictionary(Of String, Object)
        Catch ex As Exception
            Return New Dictionary(Of String, Object)
        End Try
    End Function
End Class