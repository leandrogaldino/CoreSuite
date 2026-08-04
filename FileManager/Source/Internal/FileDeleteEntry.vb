''' <summary>
''' Represents a file included in a deletion plan.
''' </summary>
Friend NotInheritable Class FileDeleteEntry
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FileDeleteEntry"/> class.
    ''' </summary>
    ''' <param name="Path">
    ''' The normalized absolute path of the file.
    ''' </param>
    ''' <param name="Length">
    ''' The file size, in bytes, recorded when the deletion plan was generated.
    ''' </param>
    Public Sub New(Path As String, Length As Long)
        Me.Path = Path
        Me.Length = Length
    End Sub
    ''' <summary>
    ''' Gets the normalized absolute path of the file to delete.
    ''' </summary>
    Public ReadOnly Property Path As String
    ''' <summary>
    ''' Gets the file size, in bytes, recorded when the deletion plan was generated.
    ''' </summary>
    Public ReadOnly Property Length As Long
End Class