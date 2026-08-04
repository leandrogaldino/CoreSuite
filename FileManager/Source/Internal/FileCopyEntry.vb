''' <summary>
''' Represents a file included in a directory copy plan.
''' </summary>
Friend NotInheritable Class FileCopyEntry
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FileCopyEntry"/> class.
    ''' </summary>
    ''' <param name="SourcePath">
    ''' The normalized absolute path of the source file.
    ''' </param>
    ''' <param name="DestinationPath">
    ''' The normalized absolute path of the destination file.
    ''' </param>
    ''' <param name="Length">
    ''' The source file size, in bytes, when the copy plan was generated.
    ''' </param>
    Public Sub New(SourcePath As String, DestinationPath As String, Length As Long)
        Me.SourcePath = SourcePath
        Me.DestinationPath = DestinationPath
        Me.Length = Length
    End Sub
    ''' <summary>
    ''' Gets the normalized absolute path of the source file.
    ''' </summary>
    Public ReadOnly Property SourcePath As String
    ''' <summary>
    ''' Gets the normalized absolute path of the destination file.
    ''' </summary>
    Public ReadOnly Property DestinationPath As String
    ''' <summary>
    ''' Gets the source file size, in bytes, recorded when the copy plan was generated.
    ''' </summary>
    Public ReadOnly Property Length As Long
End Class