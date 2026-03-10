Imports System.IO

Partial Public NotInheritable Class FileManager

    Public Class CopyDirectoryInfo

        Public Sub New()
        End Sub

        Public Sub New(source As DirectoryInfo, destination As DirectoryInfo)
            Me.Source = source
            Me.Destination = destination
        End Sub

        Public Property Source As DirectoryInfo
        Public Property Destination As DirectoryInfo
    End Class

    Public Class DeleteDirectoryInfo

        Public Sub New()
        End Sub

        Public Sub New(directory As DirectoryInfo, deleteRoot As Boolean)
            Me.Directory = directory
            Me.DeleteRoot = deleteRoot
        End Sub

        Public Property Directory As DirectoryInfo
        Public Property DeleteRoot As Boolean = True
    End Class

End Class