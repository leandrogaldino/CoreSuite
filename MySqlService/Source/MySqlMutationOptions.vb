''' <summary>
''' Provides filtering, parameter, connection, transaction, and safety settings for UPDATE and DELETE operations.
''' </summary>
Public Class MySqlMutationOptions
    Inherits MySqlCommandOptions
    ''' <summary>
    ''' Gets or sets a trusted SQL expression used after the WHERE keyword. Values must be supplied through <see cref="QueryArgs"/> instead of being concatenated into this expression.
    ''' </summary>
    Public Property Where As String
    ''' <summary>
    ''' Gets or sets parameters referenced by <see cref="Where"/>.
    ''' </summary>
    Public Property QueryArgs As IDictionary(Of String, Object) = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Gets or sets whether an UPDATE or DELETE without a WHERE clause is explicitly allowed. The default is <see langword="False"/>.
    ''' </summary>
    Public Property AllowAllRows As Boolean
End Class
