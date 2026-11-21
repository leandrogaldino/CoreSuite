''' <summary>
''' Defines one structured JOIN used by <see cref="MySqlSelectOptions"/>.
''' </summary>
Public NotInheritable Class MySqlJoin
    ''' <summary>
    ''' Initializes a new join definition.
    ''' </summary>
    ''' <param name="JoinType">The type of join to perform.</param>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Alias">An optional alias for the joined table.</param>
    ''' <param name="Condition">A trusted SQL expression used after ON. It must be omitted for a cross join. Never include untrusted user input directly.</param>
    Public Sub New(JoinType As MySqlJoinType, Table As String, Optional [Alias] As String = Nothing, Optional Condition As String = Nothing)
        Me.JoinType = JoinType
        Me.Table = Table
        Me.Alias = [Alias]
        Me.Condition = Condition
    End Sub
    ''' <summary>
    ''' Gets the type of join to perform.
    ''' </summary>
    Public ReadOnly Property JoinType As MySqlJoinType
    ''' <summary>
    ''' Gets the joined table name.
    ''' </summary>
    Public ReadOnly Property Table As String
    ''' <summary>
    ''' Gets the optional joined-table alias.
    ''' </summary>
    Public ReadOnly Property [Alias] As String
    ''' <summary>
    ''' Gets the trusted SQL expression used after ON, or <see langword="Nothing"/> for a cross join.
    ''' </summary>
    Public ReadOnly Property Condition As String
End Class
