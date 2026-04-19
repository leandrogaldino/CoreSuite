Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
''' <summary>
''' Provides functionality to merge multiple files from one or more directories into a single
''' encrypted container file and to restore them later.
''' </summary>
''' <remarks>
''' <para>
''' The <see cref="FileMerger"/> class is designed to create a compact backup-like file that contains
''' multiple files along with their relative paths and metadata. The resulting file is encrypted using
''' AES (Advanced Encryption Standard) with a key derived from a user-provided password.
''' </para>
''' 
''' <para>
''' The generated file format consists of:
''' </para>
''' <list type="number">
''' <item>
''' <description>
''' A header section containing a unique signature and encryption initialization vector (IV),
''' used to validate and decrypt the file.
''' </description>
''' </item>
''' <item>
''' <description>
''' A metadata section stored in plain text, including creation date, total file count,
''' total size, file paths, and root folders.
''' </description>
''' </item>
''' <item>
''' <description>
''' An encrypted data section containing the actual file contents and their associated paths.
''' </description>
''' </item>
''' </list>
''' 
''' <para>
''' This class supports both synchronous and asynchronous operations for merging and extracting files,
''' as well as validation and metadata inspection without full extraction.
''' </para>
''' 
''' <para>
''' It is recommended to use <see cref="IsValidFile(String)"/> before attempting to read or extract a file,
''' in order to ensure that the file matches the expected format.
''' </para>
''' 
''' <para>
''' The encryption key is derived using PBKDF2 (<see cref="Rfc2898DeriveBytes"/>) with a fixed salt
''' and iteration count. Changing these parameters will break compatibility with previously generated files.
''' </para>
''' 
''' <para>
''' This class is intended for file packaging, backup, and secure transport scenarios, but it is not a
''' replacement for full-featured archival formats (e.g., ZIP) or enterprise-grade backup solutions.
''' </para>
''' </remarks>
Public Class FileMerger
    Private Const HEADER_SIGNATURE As String = "CORE_SUITE_FILE_MERGE_SIGN:315185_19219205_69125_1351875_199714"
    ''' <summary>
    ''' Merges files from multiple directories into a single encrypted output file.
    ''' </summary>
    ''' <param name="OutputFile">The destination file path where the merged content will be stored.</param>
    ''' <param name="Directories">A collection of directory paths to include in the merge.</param>
    ''' <param name="Password">The password used to derive the encryption key.</param>
    ''' <param name="Progress">Optional progress reporter (0–100).</param>
    Public Shared Sub Merge(OutputFile As String, Directories As IEnumerable(Of String), Password As String, Optional Progress As IProgress(Of Integer) = Nothing)
        Dim AllFiles As New List(Of String)
        For Each Directory In Directories
            If IO.Directory.Exists(Directory) Then AllFiles.AddRange(IO.Directory.GetFiles(Directory, "*", SearchOption.AllDirectories))
        Next
        Dim TotalFiles As Integer = AllFiles.Count
        Dim TotalSize As Long = AllFiles.Sum(Function(f) New FileInfo(f).Length)
        Dim FilePaths As New List(Of String)()
        Dim RootFolders As New HashSet(Of String)()
        For Each FilePath In AllFiles
            Dim RootDir As String = Directories.First(Function(d) FilePath.StartsWith(d, StringComparison.InvariantCultureIgnoreCase))
            Dim RootFolderName As String = Path.GetFileName(RootDir.TrimEnd(Path.DirectorySeparatorChar))
            Dim RelPath As String = Path.Combine(RootFolderName, GetRelativePath(RootDir, FilePath))
            FilePaths.Add(RelPath)
            RootFolders.Add(RootFolderName)
        Next
        Using Stream As New FileStream(OutputFile, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim Key As Byte() = DeriveKey(Password)
            Using AES As Aes = Aes.Create()
                AES.Key = Key
                AES.GenerateIV()
                WriteHeader(Stream, AES.IV)
                Using Writer As New BinaryWriter(Stream, Encoding.UTF8, True)
                    Writer.Write(DateTime.UtcNow.ToBinary())
                    Writer.Write(TotalFiles)
                    Writer.Write(TotalSize)
                    Writer.Write(FilePaths.Count)
                    For Each Path In FilePaths
                        Dim Bytes As Byte() = Encoding.UTF8.GetBytes(Path)
                        Writer.Write(Bytes.Length)
                        Writer.Write(Bytes)
                    Next Path
                    Writer.Write(RootFolders.Count)
                    For Each RootFolder In RootFolders
                        Dim Bytes As Byte() = Encoding.UTF8.GetBytes(RootFolder)
                        Writer.Write(Bytes.Length)
                        Writer.Write(Bytes)
                    Next RootFolder
                End Using
                Using CryptoStream As New CryptoStream(Stream, AES.CreateEncryptor(), CryptoStreamMode.Write)
                    For i As Integer = 0 To AllFiles.Count - 1
                        Dim FilePath As String = AllFiles(i)
                        Dim RelPath As String = FilePaths(i)
                        Dim FileInfo As New FileInfo(FilePath)
                        Dim PathBytes As Byte() = Encoding.UTF8.GetBytes(RelPath)
                        CryptoStream.Write(BitConverter.GetBytes(PathBytes.Length), 0, 4)
                        CryptoStream.Write(PathBytes, 0, PathBytes.Length)
                        CryptoStream.Write(BitConverter.GetBytes(CInt(FileInfo.Length)), 0, 4)
                        Using FsInput As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Dim Buffer(8191) As Byte
                            Dim BytesRead As Integer
                            Do
                                BytesRead = FsInput.Read(Buffer, 0, Buffer.Length)
                                If BytesRead > 0 Then CryptoStream.Write(Buffer, 0, BytesRead)
                            Loop While BytesRead > 0
                        End Using
                        If Progress IsNot Nothing AndAlso TotalFiles > 0 Then
                            Progress.Report(CInt((i + 1) * 100 / TotalFiles))
                        End If
                    Next i
                    CryptoStream.FlushFinalBlock()
                End Using
            End Using
        End Using
    End Sub
    ''' <summary>
    ''' Extracts and restores files from a merged encrypted file into a target directory.
    ''' </summary>
    ''' <param name="InputFile">The merged file to be extracted.</param>
    ''' <param name="TargetDirectory">The destination directory where files will be restored.</param>
    ''' <param name="Password">The password used to decrypt the file.</param>
    ''' <param name="Progress">Optional progress reporter (0–100).</param>
    Public Shared Sub UnMerge(InputFile As String, TargetDirectory As String, Password As String, Optional Progress As IProgress(Of Integer) = Nothing)
        Try
            Using Stream As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim IV As Byte() = ReadHeader(Stream)
                Dim Key = DeriveKey(Password)
                Dim TotalFiles As Integer
                Dim TotalSize As Long
                Dim FilePaths As New List(Of String)()
                Dim RootFolders As New List(Of String)()
                Using Reader As New BinaryReader(Stream, Encoding.UTF8, True)
                    Dim BackupDate = DateTime.FromBinary(Reader.ReadInt64())
                    TotalFiles = Reader.ReadInt32()
                    TotalSize = Reader.ReadInt64()
                    Dim PathsCount = Reader.ReadInt32()
                    For i As Integer = 1 To PathsCount
                        Dim Length As Integer = Reader.ReadInt32()
                        Dim Bytes As Byte() = Reader.ReadBytes(Length)
                        FilePaths.Add(Encoding.UTF8.GetString(Bytes))
                    Next i

                    Dim rootCount = Reader.ReadInt32()
                    For i As Integer = 1 To rootCount
                        Dim Length As Integer = Reader.ReadInt32()
                        Dim Bytes As Byte() = Reader.ReadBytes(Length)
                        RootFolders.Add(Encoding.UTF8.GetString(Bytes))
                    Next
                End Using
                Using Aes As Aes = Aes.Create()
                    Aes.Key = Key
                    Aes.IV = IV
                    Using CryptoStream As New CryptoStream(Stream, Aes.CreateDecryptor(), CryptoStreamMode.Read)
                        Using Reader As New BinaryReader(CryptoStream, Encoding.UTF8, True)
                            For i As Integer = 0 To TotalFiles - 1
                                Dim PathLength As Integer = Reader.ReadInt32()
                                Dim PathBytes As Byte() = Reader.ReadBytes(PathLength)
                                Dim RelPath As String = Encoding.UTF8.GetString(PathBytes)
                                Dim FileLength As Integer = Reader.ReadInt32()
                                Dim FileBytes(FileLength - 1) As Byte
                                Dim BytesRead As Integer = 0
                                While BytesRead < FileLength
                                    BytesRead += CryptoStream.Read(FileBytes, BytesRead, FileLength - BytesRead)
                                End While
                                Dim OutPath As String = Path.Combine(TargetDirectory, RelPath)
                                Directory.CreateDirectory(Path.GetDirectoryName(OutPath))
                                Using FsOut As New FileStream(OutPath, FileMode.Create, FileAccess.Write, FileShare.None)
                                    FsOut.Write(FileBytes, 0, FileBytes.Length)
                                End Using
                                If Progress IsNot Nothing AndAlso TotalFiles > 0 Then
                                    Progress.Report(CInt((i + 1) * 100 / TotalFiles))
                                End If
                            Next
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As CryptographicException
            Throw New UnauthorizedAccessException("Incorrect password or invalid backup file.")
        End Try
    End Sub
    ''' <summary>
    ''' Determines whether the specified file is a valid merged file generated by this class.
    ''' </summary>
    ''' <param name="FilePath">The file path to validate.</param>
    ''' <returns>True if the file is valid; otherwise, False.</returns>
    Public Shared Function IsValidFile(FilePath As String) As Boolean
        Try
            Dim Info As New FileInfo(FilePath)
            If Not Info.Exists OrElse Info.Length < 10 Then Return False
            Using Stream As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using Reader As New BinaryReader(Stream, Encoding.UTF8, True)
                    Dim Signature As String = Nothing
                    If Not TryReadString(Reader, Stream, Signature) Then
                        Return False
                    End If
                    If Signature <> HEADER_SIGNATURE Then
                        Return False
                    End If
                    If Stream.Length - Stream.Position < 4 Then Return False
                    Dim IVLength As Integer = Reader.ReadInt32()
                    If IVLength <= 0 OrElse IVLength > 64 Then Return False
                    If Stream.Length - Stream.Position < IVLength Then Return False
                    Dim IV As Byte() = Reader.ReadBytes(IVLength)
                    If IV.Length <> IVLength Then Return False
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function
    Private Shared Function TryReadString(Reader As BinaryReader, Stream As FileStream, ByRef Result As String) As Boolean
        Result = Nothing
        Dim Length As Integer = 0
        Dim Shift As Integer = 0
        While True
            If Stream.Position >= Stream.Length Then Return False
            Dim ReadedByte As Byte = Reader.ReadByte()
            Length = Length Or ((ReadedByte And &H7F) << Shift)
            Shift += 7
            If Shift > 35 Then Return False
            If (ReadedByte And &H80) = 0 Then Exit While
        End While
        If Length < 0 OrElse Length > Stream.Length - Stream.Position Then
            Return False
        End If
        Dim Bytes As Byte() = Reader.ReadBytes(Length)
        If Bytes.Length <> Length Then Return False
        Result = Encoding.UTF8.GetString(Bytes)
        Return True
    End Function
    ''' <summary>
    ''' Asynchronously merges files from multiple directories into a single encrypted output file.
    ''' </summary>
    ''' <param name="OutputFile">The destination file path where the merged content will be stored.</param>
    ''' <param name="Directories">A collection of directory paths to include in the merge.</param>
    ''' <param name="password">The password used to derive the encryption key.</param>
    ''' <param name="progress">Optional progress reporter (0–100).</param>
    ''' <returns>A task that represents the asynchronous operation.</returns>
    Public Shared Async Function MergeAsync(OutputFile As String, Directories As IEnumerable(Of String), password As String, Optional progress As IProgress(Of Integer) = Nothing) As Task
        Await Task.Run(Sub() Merge(OutputFile, Directories, password, progress))
    End Function
    ''' <summary>
    ''' Asynchronously extracts and restores files from a merged encrypted file.
    ''' </summary>
    ''' <param name="InputFile">The merged file to be extracted.</param>
    ''' <param name="TargetDirectory">The destination directory where files will be restored.</param>
    ''' <param name="Password">The password used to decrypt the file.</param>
    ''' <param name="Progress">Optional progress reporter (0–100).</param>
    ''' <returns>A task that represents the asynchronous operation.</returns>
    Public Shared Async Function UnMergeAsync(InputFile As String, TargetDirectory As String, Password As String, Optional Progress As IProgress(Of Integer) = Nothing) As Task
        Await Task.Run(Sub() UnMerge(InputFile, TargetDirectory, Password, Progress))
    End Function
    ''' <summary>
    ''' Retrieves metadata information from a merged file without extracting its contents.
    ''' </summary>
    ''' <param name="InputFile">The merged file to read metadata from.</param>
    ''' <returns>A dictionary containing metadata such as creation date, total files, total size, file paths, and root folders.</returns>
    Public Shared Function GetMetadata(InputFile As String) As Dictionary(Of String, Object)
        Dim Metadata As New Dictionary(Of String, Object)
        Using Stream As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read)
            Dim IV As Byte() = ReadHeader(Stream)
            Using Reader As New BinaryReader(Stream, Encoding.UTF8, True)
                Dim BackupDate As DateTime = DateTime.FromBinary(Reader.ReadInt64()).ToLocalTime()
                Dim TotalFiles As Integer = Reader.ReadInt32()
                Dim TotalSize As Long = Reader.ReadInt64()
                Dim PathsCount As Integer = Reader.ReadInt32()
                Dim FilePaths As New List(Of String)()
                For i As Integer = 1 To PathsCount
                    Dim Len As Integer = Reader.ReadInt32()
                    Dim Bytes As Byte() = Reader.ReadBytes(Len)
                    FilePaths.Add(Encoding.UTF8.GetString(Bytes))
                Next i
                Dim RootCount As Integer = Reader.ReadInt32()
                Dim RootFolders As New List(Of String)()
                For i As Integer = 1 To RootCount
                    Dim Len As Integer = Reader.ReadInt32()
                    Dim Bytes As Byte() = Reader.ReadBytes(Len)
                    RootFolders.Add(Encoding.UTF8.GetString(Bytes))
                Next i
                Metadata("CreationDate") = BackupDate
                Metadata("TotalFiles") = TotalFiles
                Metadata("TotalSize") = TotalSize
                Metadata("FilePaths") = FilePaths
                Metadata("RootFolders") = RootFolders
            End Using
        End Using
        Return Metadata
    End Function
    Private Shared Function DeriveKey(Password As String) As Byte()
        Dim Pdb As New Rfc2898DeriveBytes(Password, Encoding.UTF8.GetBytes("fd529abe-3283-487e-b6fd-127d9868c658"), 10000)
        Return Pdb.GetBytes(32)
    End Function
    Private Shared Sub WriteHeader(Stream As FileStream, IV As Byte())
        Using Writer As New BinaryWriter(Stream, Encoding.UTF8, True)
            Writer.Write(HEADER_SIGNATURE)
            Writer.Write(IV.Length)
            Writer.Write(IV)
        End Using
    End Sub
    Private Shared Function ReadHeader(Stream As FileStream) As Byte()
        Using Reader As New BinaryReader(Stream, Encoding.UTF8, True)
            Dim Signature As String = Reader.ReadString()
            If Signature <> HEADER_SIGNATURE Then Throw New InvalidDataException("Invalid file.")
            Dim IVLength = Reader.ReadInt32()
            Return Reader.ReadBytes(IVLength)
        End Using
    End Function
    Private Shared Function GetRelativePath(BasePath As String, FullPath As String) As String
        Dim BaseUri = New Uri(If(BasePath.EndsWith(Path.DirectorySeparatorChar), BasePath, BasePath & Path.DirectorySeparatorChar))
        Dim FullUri = New Uri(FullPath)
        Return Uri.UnescapeDataString(BaseUri.MakeRelativeUri(FullUri).ToString().Replace("/"c, Path.DirectorySeparatorChar))
    End Function
End Class