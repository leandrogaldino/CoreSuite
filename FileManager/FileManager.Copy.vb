Imports System.IO

Partial Public NotInheritable Class FileManager

#Region "Copy File"

    Public Shared Async Function CopyFileAsync(
        source As FileInfo,
        destination As FileInfo,
        Optional progress As IProgress(Of Integer) = Nothing
    ) As Task

        Dim bufferSize As Integer = 1024 * 1024
        Dim totalSize = source.Length
        Dim handledSize As Long = 0

        Using sourceStream As New FileStream(source.FullName, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, True)
            Using destStream As New FileStream(destination.FullName, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, True)

                Dim buffer(bufferSize - 1) As Byte
                Dim bytesRead As Integer

                Do
                    bytesRead = Await sourceStream.ReadAsync(buffer, 0, buffer.Length)
                    If bytesRead > 0 Then
                        Await destStream.WriteAsync(buffer, 0, bytesRead)
                        handledSize += bytesRead
                        progress?.Report(CalcPercent(handledSize, totalSize))
                    End If
                Loop While bytesRead > 0

            End Using
        End Using
    End Function

#End Region

#Region "Copy Directories"

    Public Shared Async Function CopyDirectoriesAsync(
        directories As List(Of CopyDirectoryInfo),
        Optional progress As IProgress(Of Integer) = Nothing
    ) As Task

        Dim totalSize As Long = 0
        For Each info In directories
            totalSize += GetDirectorySize(info.Source)
        Next

        Dim handledSize As Long = 0

        For Each info In directories
            Dim copied = Await CopyDirectoryAsync(info, totalSize, handledSize, progress)
            handledSize += copied
        Next
    End Function

    Private Shared Async Function CopyDirectoryAsync(
        info As CopyDirectoryInfo,
        totalSize As Long,
        currentHandled As Long,
        Optional progress As IProgress(Of Integer) = Nothing
    ) As Task(Of Long)

        If Not info.Source.Exists Then
            Throw New DirectoryNotFoundException(
                $"Source directory '{info.Source.FullName}' not found.")
        End If

        If Not info.Destination.Exists Then
            info.Destination.Create()
        End If

        Dim copiedSize As Long = 0
        Dim files = info.Source.GetFiles("*.*", SearchOption.AllDirectories)

        For Each file In files
            Dim destPath = Path.Combine(
                info.Destination.FullName,
                file.FullName.Substring(info.Source.FullName.Length + 1))

            Directory.CreateDirectory(Path.GetDirectoryName(destPath))

            Using src As New FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192, True)
                Using dst As New FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, True)

                    Dim buffer(8191) As Byte
                    Dim bytesRead As Integer

                    Do
                        bytesRead = Await src.ReadAsync(buffer, 0, buffer.Length)
                        If bytesRead > 0 Then
                            Await dst.WriteAsync(buffer, 0, bytesRead)
                            copiedSize += bytesRead
                            progress?.Report(
                                CalcPercent(currentHandled + copiedSize, totalSize))
                        End If
                    Loop While bytesRead > 0

                End Using
            End Using
        Next

        Return copiedSize
    End Function

#End Region

End Class
