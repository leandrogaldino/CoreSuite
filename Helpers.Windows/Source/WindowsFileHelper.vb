Imports System.IO
Imports System.Runtime.InteropServices

''' <summary>Provides Windows file movement with replacement semantics.</summary>
Public NotInheritable Class WindowsFileHelper
    Private Sub New()
    End Sub
    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="MoveFileExW")>
    Private Shared Function MoveFileEx(Source As String, Destination As String, Flags As UInteger) As Boolean
    End Function

    ''' <summary>Moves a file, optionally replacing an existing destination without deleting it first.</summary>
    Public Shared Sub MoveFile(Source As String, Destination As String, Overwrite As Boolean)
        If Not Overwrite Then
            File.Move(Source, Destination)
            Return
        End If
        If Source Is Nothing Then Throw New ArgumentNullException(NameOf(Source))
        If Destination Is Nothing Then Throw New ArgumentNullException(NameOf(Destination))
        Dim FullSource = Path.GetFullPath(Source)
        Dim FullDestination = Path.GetFullPath(Destination)
        ' MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED, matching Windows File.Move.
        If Not MoveFileEx(FullSource, FullDestination, 3UI) Then
            Dim ErrorCode = Marshal.GetLastWin32Error()
            Select Case ErrorCode
                Case 2 : Throw New FileNotFoundException("Source file was not found.", FullSource)
                Case 3 : Throw New DirectoryNotFoundException("The source or destination directory was not found.")
                Case 5 : Throw New UnauthorizedAccessException("Access to the file path is denied.")
                Case Else
                    Throw New IOException(New System.ComponentModel.Win32Exception(ErrorCode).Message, (-2147024896 Or ErrorCode))
            End Select
        End If
    End Sub

End Class
