''' <summary>
''' Provides projection, filtering, ordering, paging, connection, and transaction settings for a SELECT operation.
''' </summary>
Public Class MySqlSelectOptions
    Inherits MySqlCommandOptions
    ''' <summary>
    ''' Gets the safely quoted column names to return. When both this collection and <see cref="TrustedExpressions"/> are empty, all columns are returned.
    ''' </summary>
    Public ReadOnly Property Columns As IList(Of String) = New List(Of String)()
    ''' <summary>
    ''' Gets trusted raw SQL expressions to include in the projection, such as aggregate expressions or explicit aliases. Never add untrusted user input to this collection.
    ''' </summary>
    Public ReadOnly Property TrustedExpressions As IList(Of String) = New List(Of String)()
    ''' <summary>
    ''' Gets or sets a trusted SQL expression used after the WHERE keyword. Values must be supplied through <see cref="QueryArgs"/>.
    ''' </summary>
    Public Property Where As String
    ''' <summary>
    ''' Gets or sets parameters referenced by <see cref="Where"/>.
    ''' </summary>
    Public Property QueryArgs As IDictionary(Of String, Object) = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Gets the safely quoted columns used by the ORDER BY clause.
    ''' </summary>
    Public ReadOnly Property OrderBy As IList(Of MySqlOrderBy) = New List(Of MySqlOrderBy)()
    ''' <summary>
    ''' Gets or sets whether the SELECT operation uses DISTINCT.
    ''' </summary>
    Public Property Distinct As Boolean
    ''' <summary>
    ''' Gets or sets the maximum number of rows to return. A null value does not apply a limit.
    ''' </summary>
    Public Property Limit As Integer?
    ''' <summary>
    ''' Gets or sets the number of rows to skip. A null value or zero does not apply an offset.
    ''' </summary>
    Public Property Offset As Integer?
End Class
