Imports System.IO
''' <summary>
''' Describes a directory deletion request and whether its root directory should also be removed.
''' </summary>
Public NotInheritable Class DeleteDirectoryInfo
    ''' <summary>
    ''' Initializes a new empty instance of the <see cref="DeleteDirectoryInfo"/> class.
    ''' </summary>
    ''' <remarks>
    ''' The <see cref="Directory"/> property must be assigned before the instance is used in a deletion operation. <see cref="DeleteRoot"/> defaults to <see langword="True"/>.
    ''' </remarks>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DeleteDirectoryInfo"/> class.
    ''' </summary>
    ''' <param name="Directory">
    ''' The directory whose contents will be deleted.
    ''' </param>
    ''' <param name="DeleteRoot">
    ''' <see langword="True"/> to delete the root directory after its contents are removed; otherwise, <see langword="False"/>.
    ''' </param>
    Public Sub New(Directory As DirectoryInfo, DeleteRoot As Boolean)
        Me.Directory = Directory
        Me.DeleteRoot = DeleteRoot
    End Sub
    ''' <summary>
    ''' Gets or sets the root directory represented by the deletion request.
    ''' </summary>
    ''' <value>
    ''' A <see cref="DirectoryInfo"/> representing the directory to process.
    ''' </value>
    Public Property Directory As DirectoryInfo
    ''' <summary>
    ''' Gets or sets whether the root directory should be deleted after its contents are removed.
    ''' </summary>
    ''' <value>
    ''' <see langword="True"/> to delete the root directory; otherwise, <see langword="False"/>. The default is <see langword="True"/>.
    ''' </value>
    Public Property DeleteRoot As Boolean = True
End Class