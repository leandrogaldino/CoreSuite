Imports System.IO
''' <summary>
''' Provides helper methods for file system operations such as
''' checking file locks and safely deleting files or directories.
''' </summary>
Public Class FileHelper
    ''' <summary>
    ''' Attempts to delete a file if it is not currently in use by another process.
    ''' </summary>
    ''' <param name="File">
    ''' The <see cref="FileInfo"/> representing the file to be deleted.
    ''' </param>
    ''' <returns>
    ''' <c>True</c> if the file is not locked and was successfully deleted;
    ''' <c>False</c> if the file is in use or could not be deleted.
    ''' </returns>
    Public Shared Function TryDeleteFile(File As FileInfo) As Boolean
        Dim Deleted As Boolean
        Try
            Dim Stream As FileStream = IO.File.Open(File.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
            Stream.Close()
            File.Delete()
            Deleted = True
        Catch ex As IOException
            Deleted = False
        End Try
        Return Deleted
    End Function
    ''' <summary>
    ''' Attempts to delete a directory and its contents.
    ''' The operation fails if any file within the directory is locked or in use.
    ''' </summary>
    ''' <param name="Directory">
    ''' The <see cref="DirectoryInfo"/> representing the directory to be deleted.
    ''' </param>
    ''' <returns>
    ''' <c>True</c> if the directory was successfully deleted;
    ''' <c>False</c> if the deletion failed due to locked files or I/O errors.
    ''' </returns>
    Public Shared Function TryDeleteDirectory(Directory As DirectoryInfo) As Boolean
        Dim Deleted As Boolean
        Try
            Directory.Delete(True)
            Deleted = True
        Catch ex As IOException
            Deleted = False
        End Try
        Return Deleted
    End Function
    ''' <summary>
    ''' Determines whether a file is currently locked by another process.
    ''' </summary>
    ''' <param name="File">
    ''' The <see cref="FileInfo"/> representing the file to check.
    ''' </param>
    ''' <returns>
    ''' <c>True</c> if the file is locked or inaccessible;
    ''' <c>False</c> if the file is available for exclusive access.
    ''' </returns>
    Public Shared Function IsFileLocked(File As FileInfo) As Boolean
        Try
            Using Stream As FileStream = File.Open(FileMode.Open, FileAccess.Read, FileShare.None)
                Stream.Close()
            End Using
        Catch ex As IOException
            Return True
        End Try
        Return False
    End Function
End Class
