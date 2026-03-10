Imports System.IO

Partial Public NotInheritable Class FileManager

#Region "Delete Files"

    Public Shared Async Function DeleteFilesAsync(
        files As List(Of FileInfo),
        Optional progress As IProgress(Of Integer) = Nothing
    ) As Task

        Dim totalSize = files.Sum(Function(f) f.Length)
        Dim handledSize As Long = 0

        For Each file In files
            Await Task.Run(Sub() file.Delete())
            handledSize += file.Length
            progress?.Report(CalcPercent(handledSize, totalSize))
        Next
    End Function

#End Region

#Region "Delete Directories"

    Public Shared Async Function DeleteDirectoriesAsync(
        directories As List(Of DeleteDirectoryInfo),
        Optional progress As IProgress(Of Integer) = Nothing
    ) As Task

        Dim allFiles As New List(Of FileInfo)
        Dim allDirs As New List(Of DirectoryInfo)

        For Each info In directories
            If info.Directory.Exists Then
                allFiles.AddRange(info.Directory.GetFiles("*", SearchOption.AllDirectories))
                allDirs.AddRange(info.Directory.GetDirectories("*", SearchOption.AllDirectories))
            End If
        Next

        allFiles = allFiles.
            OrderByDescending(Function(f) f.FullName.Count(Function(c) c = Path.DirectorySeparatorChar)).
            ToList()

        Dim totalSize = allFiles.Sum(Function(f) f.Length)
        Dim handledSize As Long = 0

        For Each file In allFiles
            Await Task.Run(Sub() file.Delete())
            handledSize += file.Length
            progress?.Report(CalcPercent(handledSize, totalSize))
        Next

        For Each Directory In allDirs.
            OrderByDescending(Function(d) d.FullName.Count(Function(c) c = Path.DirectorySeparatorChar))
            Await Task.Run(Sub() Directory.Delete())
        Next

        For Each info In directories
            If info.DeleteRoot AndAlso info.Directory.Exists Then
                Await Task.Run(Sub() info.Directory.Delete(True))
            End If
        Next
    End Function

#End Region

#Region "Delete Directory Content"

    Public Shared Async Function DeleteDirectoryContentAsync(
        directory As DirectoryInfo,
        exceptDirectories As List(Of DirectoryInfo),
        exceptFiles As List(Of FileInfo)
    ) As Task

        For Each file In directory.GetFiles()
            If exceptFiles Is Nothing OrElse Not exceptFiles.Contains(file) Then
                Await Task.Run(Sub() file.Delete())
            End If
        Next

        For Each subDir In directory.GetDirectories()
            If exceptDirectories Is Nothing OrElse Not exceptDirectories.Contains(subDir) Then
                Await Task.Run(Sub() subDir.Delete(True))
            End If
        Next
    End Function

#End Region

End Class