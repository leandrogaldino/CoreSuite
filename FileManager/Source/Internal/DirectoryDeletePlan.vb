''' <summary>
''' Contains the roots, files, and directories required by a directory deletion operation.
''' </summary>
Friend NotInheritable Class DirectoryDeletePlan
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DirectoryDeletePlan"/> class.
    ''' </summary>
    ''' <param name="Roots">
    ''' The root directories represented by the deletion operation.
    ''' </param>
    ''' <param name="Files">
    ''' The files that must be deleted.
    ''' </param>
    ''' <param name="Directories">
    ''' The child directories that must be deleted after their files are removed.
    ''' </param>
    Public Sub New(Roots As List(Of DirectoryDeleteRoot), Files As List(Of FileDeleteEntry), Directories As List(Of String))
        Me.Roots = Roots
        Me.Files = Files
        Me.Directories = Directories
    End Sub
    ''' <summary>
    ''' Gets the root directories represented by the deletion operation.
    ''' </summary>
    Public ReadOnly Property Roots As List(Of DirectoryDeleteRoot)
    ''' <summary>
    ''' Gets the files included in the deletion operation.
    ''' </summary>
    Public ReadOnly Property Files As List(Of FileDeleteEntry)
    ''' <summary>
    ''' Gets the child directory paths included in the deletion operation.
    ''' </summary>
    Public ReadOnly Property Directories As List(Of String)
    ''' <summary>
    ''' Gets the combined size, in bytes, of all files included in the deletion operation.
    ''' </summary>
    Public ReadOnly Property TotalSize As Long
        Get
            Return Files.Sum(Function(CurrentFile) CurrentFile.Length)
        End Get
    End Property
End Class