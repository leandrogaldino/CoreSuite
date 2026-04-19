Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class FileMerger
    Private Const HEADER_SIGNATURE As String = "FILEMERGESIGN"

    ' ================== CRIAR BACKUP SÍNCRONO ==================
    Public Shared Sub Merge(OutputFile As String, Directories As IEnumerable(Of String), Password As String, Optional Progress As IProgress(Of Integer) = Nothing)
        ' Lista todos os arquivos
        Dim allFiles As New List(Of String)
        For Each d In Directories
            If Directory.Exists(d) Then allFiles.AddRange(Directory.GetFiles(d, "*", SearchOption.AllDirectories))
        Next

        Dim totalFiles = allFiles.Count
        Dim totalSize As Long = allFiles.Sum(Function(f) New FileInfo(f).Length)
        Dim filePaths As New List(Of String)()
        Dim rootFolders As New HashSet(Of String)()

        ' Caminhos relativos e rootFolders
        For Each filePath In allFiles
            Dim rootDir = Directories.First(Function(d) filePath.StartsWith(d, StringComparison.InvariantCultureIgnoreCase))
            Dim rootFolderName = Path.GetFileName(rootDir.TrimEnd(Path.DirectorySeparatorChar))
            Dim relPath = Path.Combine(rootFolderName, GetRelativePath(rootDir, filePath))
            filePaths.Add(relPath)
            rootFolders.Add(rootFolderName)
        Next

        Using fs As New FileStream(OutputFile, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim key = DeriveKey(Password)
            Using aes As Aes = Aes.Create()
                aes.Key = key
                aes.GenerateIV()
                ' Escreve header
                WriteHeader(fs, aes.IV)

                ' --- Grava metadados em claro ---
                Using bw As New BinaryWriter(fs, Encoding.UTF8, True)
                    bw.Write(DateTime.UtcNow.ToBinary())
                    bw.Write(totalFiles)
                    bw.Write(totalSize)

                    bw.Write(filePaths.Count)
                    For Each path In filePaths
                        Dim bytes = Encoding.UTF8.GetBytes(path)
                        bw.Write(bytes.Length)
                        bw.Write(bytes)
                    Next

                    bw.Write(rootFolders.Count)
                    For Each rf In rootFolders
                        Dim bytes = Encoding.UTF8.GetBytes(rf)
                        bw.Write(bytes.Length)
                        bw.Write(bytes)
                    Next
                End Using

                ' --- Criptografa e grava os arquivos ---
                Using cryptoStream As New CryptoStream(fs, aes.CreateEncryptor(), CryptoStreamMode.Write)
                    For i As Integer = 0 To allFiles.Count - 1
                        Dim filePath = allFiles(i)
                        Dim relPath = filePaths(i)
                        Dim fileInfo As New FileInfo(filePath)

                        ' Caminho relativo
                        Dim pathBytes = Encoding.UTF8.GetBytes(relPath)
                        cryptoStream.Write(BitConverter.GetBytes(pathBytes.Length), 0, 4)
                        cryptoStream.Write(pathBytes, 0, pathBytes.Length)

                        ' Conteúdo do arquivo
                        cryptoStream.Write(BitConverter.GetBytes(CInt(fileInfo.Length)), 0, 4)

                        Using fsInput As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Dim buffer(8191) As Byte
                            Dim bytesRead As Integer
                            Do
                                bytesRead = fsInput.Read(buffer, 0, buffer.Length)
                                If bytesRead > 0 Then cryptoStream.Write(buffer, 0, bytesRead)
                            Loop While bytesRead > 0
                        End Using

                        If Progress IsNot Nothing AndAlso totalFiles > 0 Then
                            Progress.Report(CInt((i + 1) * 100 / totalFiles))
                        End If
                    Next
                    cryptoStream.FlushFinalBlock()
                End Using
            End Using
        End Using
    End Sub

    ' ================== RESTAURAR BACKUP SÍNCRONO ==================
    Public Shared Sub UnMerge(InputFile As String, TargetDirectory As String, Password As String, Optional Progress As IProgress(Of Integer) = Nothing)
        Try
            Using fs As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read)
                ' Lê o header
                Dim iv As Byte() = ReadHeader(fs)
                Dim key = DeriveKey(Password)

                ' --- Lê metadados em claro ---
                Dim totalFiles As Integer
                Dim totalSize As Long
                Dim filePaths As New List(Of String)()
                Dim rootFolders As New List(Of String)()

                Using br As New BinaryReader(fs, Encoding.UTF8, True)
                    Dim backupDate = DateTime.FromBinary(br.ReadInt64())
                    totalFiles = br.ReadInt32()
                    totalSize = br.ReadInt64()

                    Dim pathsCount = br.ReadInt32()
                    For i As Integer = 1 To pathsCount
                        Dim length = br.ReadInt32()
                        Dim bytes = br.ReadBytes(length)
                        filePaths.Add(Encoding.UTF8.GetString(bytes))
                    Next

                    Dim rootCount = br.ReadInt32()
                    For i As Integer = 1 To rootCount
                        Dim length = br.ReadInt32()
                        Dim bytes = br.ReadBytes(length)
                        rootFolders.Add(Encoding.UTF8.GetString(bytes))
                    Next
                End Using

                ' --- Descriptografa e restaura arquivos ---
                Using aes As Aes = Aes.Create()
                    aes.Key = key
                    aes.IV = iv

                    Using cryptoStream As New CryptoStream(fs, aes.CreateDecryptor(), CryptoStreamMode.Read)
                        Using br As New BinaryReader(cryptoStream, Encoding.UTF8, True)
                            For i As Integer = 0 To totalFiles - 1
                                ' Lê caminho relativo
                                Dim pathLength = br.ReadInt32()
                                Dim pathBytes = br.ReadBytes(pathLength)
                                Dim relPath = Encoding.UTF8.GetString(pathBytes)

                                ' Lê conteúdo do arquivo
                                Dim fileLength = br.ReadInt32()
                                Dim fileBytes(fileLength - 1) As Byte
                                Dim bytesRead As Integer = 0
                                While bytesRead < fileLength
                                    bytesRead += cryptoStream.Read(fileBytes, bytesRead, fileLength - bytesRead)
                                End While

                                ' Grava no diretório de destino
                                Dim outPath = Path.Combine(TargetDirectory, relPath)
                                Directory.CreateDirectory(Path.GetDirectoryName(outPath))
                                Using fsOut As New FileStream(outPath, FileMode.Create, FileAccess.Write, FileShare.None)
                                    fsOut.Write(fileBytes, 0, fileBytes.Length)
                                End Using

                                If Progress IsNot Nothing AndAlso totalFiles > 0 Then
                                    Progress.Report(CInt((i + 1) * 100 / totalFiles))
                                End If
                            Next
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As CryptographicException
            Throw New UnauthorizedAccessException("Senha incorreta ou arquivo de backup inválido.")
        End Try
    End Sub

    Public Shared Function IsValidFile(filePath As String) As Boolean
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using br As New BinaryReader(fs, Encoding.UTF8, True)
                    ' Lê a assinatura
                    Dim signature = br.ReadString()
                    If signature <> HEADER_SIGNATURE Then
                        Return False
                    End If

                    ' Lê tamanho do IV e valida se o arquivo tem bytes suficientes
                    Dim ivLength = br.ReadInt32()
                    If ivLength <= 0 OrElse ivLength > 64 Then
                        ' 64 é um limite arbitrário de segurança (IV normalmente é 16 bytes no AES)
                        Return False
                    End If
                    Dim iv = br.ReadBytes(ivLength)
                    If iv.Length <> ivLength Then
                        Return False
                    End If

                    Return True
                End Using
            End Using
        Catch
            ' Se der qualquer erro (arquivo inexistente, corrompido, sem permissão, etc)
            Return False
        End Try
    End Function

    ' ================== MÉTODO PARA EXECUTAR EM BACKGROUND ==================
    Public Shared Async Function MergeAsync(OutputFile As String, Directories As IEnumerable(Of String), password As String, Optional progress As IProgress(Of Integer) = Nothing) As Task
        ' Roda o Merge síncrono em outra thread e espera
        Await Task.Run(Sub() Merge(OutputFile, Directories, password, progress))
    End Function

    Public Shared Async Function UnMergeAsync(InputFile As String, TargetDirectory As String, Password As String, Optional Progress As IProgress(Of Integer) = Nothing) As Task
        ' Roda o Unmerge síncrono em outra thread e espera
        Await Task.Run(Sub() UnMerge(InputFile, TargetDirectory, Password, Progress))
    End Function

    ' ================== LER METADADOS ==================
    Public Shared Function GetMetadata(InputFile As String) As Dictionary(Of String, Object)
        Dim metadata As New Dictionary(Of String, Object)
        Using fs As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read)
            ' Header
            Dim iv = ReadHeader(fs)

            Using br As New BinaryReader(fs, Encoding.UTF8, True)
                Dim backupDate = DateTime.FromBinary(br.ReadInt64()).ToLocalTime()
                Dim totalFiles = br.ReadInt32()
                Dim totalSize = br.ReadInt64()

                ' FilePaths
                Dim pathsCount = br.ReadInt32()
                Dim filePaths As New List(Of String)()
                For i As Integer = 1 To pathsCount
                    Dim len = br.ReadInt32()
                    Dim bytes = br.ReadBytes(len)
                    filePaths.Add(Encoding.UTF8.GetString(bytes))
                Next

                ' RootFolders
                Dim rootCount = br.ReadInt32()
                Dim rootFolders As New List(Of String)()
                For i As Integer = 1 To rootCount
                    Dim len = br.ReadInt32()
                    Dim bytes = br.ReadBytes(len)
                    rootFolders.Add(Encoding.UTF8.GetString(bytes))
                Next

                metadata("CreationDate") = backupDate
                metadata("TotalFiles") = totalFiles
                metadata("TotalSize") = totalSize
                metadata("FilePaths") = filePaths
                metadata("RootFolders") = rootFolders
            End Using
        End Using
        Return metadata
    End Function

    ' ================== SUPORTE ==================
    Private Shared Function DeriveKey(password As String) As Byte()
        Dim pdb = New Rfc2898DeriveBytes(password, Encoding.UTF8.GetBytes("fd529abe-3283-487e-b6fd-127d9868c658"), 10000)
        Return pdb.GetBytes(32)
    End Function

    Private Shared Sub WriteHeader(fs As FileStream, iv As Byte())
        Using bw As New BinaryWriter(fs, Encoding.UTF8, True)
            bw.Write(HEADER_SIGNATURE)
            bw.Write(iv.Length)
            bw.Write(iv)
        End Using
    End Sub

    Private Shared Function ReadHeader(fs As FileStream) As Byte()
        Using br As New BinaryReader(fs, Encoding.UTF8, True)
            Dim signature = br.ReadString()
            If signature <> HEADER_SIGNATURE Then Throw New InvalidDataException("Arquivo inválido.")
            Dim ivLength = br.ReadInt32()
            Return br.ReadBytes(ivLength)
        End Using
    End Function

    Private Shared Function GetRelativePath(basePath As String, fullPath As String) As String
        Dim baseUri = New Uri(If(basePath.EndsWith(Path.DirectorySeparatorChar), basePath, basePath & Path.DirectorySeparatorChar))
        Dim fullUri = New Uri(fullPath)
        Return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fullUri).ToString().Replace("/"c, Path.DirectorySeparatorChar))
    End Function

End Class