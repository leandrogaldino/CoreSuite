''' <summary>
''' Contains all information required to copy one directory tree.
''' </summary>
Friend NotInheritable Class DirectoryCopyPlan
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DirectoryCopyPlan"/> class.
    ''' </summary>
    ''' <param name="DestinationRoot">
    ''' The normalized absolute path of the destination root directory.
    ''' </param>
    ''' <param name="Directories">
    ''' The destination directories that must be created.
    ''' </param>
    ''' <param name="Files">
    ''' The source and destination file mappings that must be copied.
    ''' </param>
    Public Sub New(DestinationRoot As String, Directories As List(Of String), Files As List(Of FileCopyEntry))
        Me.DestinationRoot = DestinationRoot
        Me.Directories = Directories
        Me.Files = Files
    End Sub
    ''' <summary>
    ''' Gets the normalized absolute path of the destination root directory.
    ''' </summary>
    Public ReadOnly Property DestinationRoot As String
    ''' <summary>
    ''' Gets the destination directory paths that must be created.
    ''' </summary>
    Public ReadOnly Property Directories As List(Of String)
    ''' <summary>
    ''' Gets the file mappings included in the copy plan.
    ''' </summary>
    Public ReadOnly Property Files As List(Of FileCopyEntry)
    ''' <summary>
    ''' Gets the combined size, in bytes, of all source files included in the copy plan.
    ''' </summary>
    Public ReadOnly Property TotalSize As Long
        Get
            Return Files.Sum(Function(CurrentFile) CurrentFile.Length)
        End Get
    End Property
End Class