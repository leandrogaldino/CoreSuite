''' <summary>
''' Represents a safely quoted ORDER BY column and its direction.
''' </summary>
Public NotInheritable Class MySqlOrderBy
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySqlOrderBy"/> class.
    ''' </summary>
    ''' <param name="column">The column name, optionally qualified by a table or schema.</param>
    ''' <param name="direction">The sort direction.</param>
    Public Sub New(column As String, Optional direction As MySqlSortDirection = MySqlSortDirection.Ascending)
        If Not [Enum].IsDefined(GetType(MySqlSortDirection), direction) Then Throw New ArgumentOutOfRangeException(NameOf(direction), direction, "The sort direction is not valid.")
        Me.Column = RequireValue(column, NameOf(column))
        Me.Direction = direction
    End Sub
    ''' <summary>
    ''' Gets the column name.
    ''' </summary>
    Public ReadOnly Property Column As String
    ''' <summary>
    ''' Gets the sort direction.
    ''' </summary>
    Public ReadOnly Property Direction As MySqlSortDirection
End Class
