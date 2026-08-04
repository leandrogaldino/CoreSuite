''' <summary>
''' Represents a root directory included in a directory deletion plan.
''' </summary>
Friend NotInheritable Class DirectoryDeleteRoot
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DirectoryDeleteRoot"/> class.
    ''' </summary>
    ''' <param name="Path">
    ''' The normalized absolute path of the root directory.
    ''' </param>
    ''' <param name="DeleteRoot">
    ''' <see langword="True"/> to delete the root after its contents are removed; otherwise, <see langword="False"/>.
    ''' </param>
    Public Sub New(Path As String, DeleteRoot As Boolean)
        Me.Path = Path
        Me.DeleteRoot = DeleteRoot
    End Sub
    ''' <summary>
    ''' Gets the normalized absolute path of the root directory.
    ''' </summary>
    Public ReadOnly Property Path As String
    ''' <summary>
    ''' Gets whether the root directory should be deleted after its contents are removed.
    ''' </summary>
    Public ReadOnly Property DeleteRoot As Boolean
End Class