Imports MySql.Data.MySqlClient
''' <summary>
''' Provides connection, transaction, and timeout settings for a MySQL command.
''' </summary>
Public Class MySqlCommandOptions
    ''' <summary>
    ''' Gets or sets an optional connection to use. Closed external connections are opened for the command and closed again afterward; open external connections remain open.
    ''' </summary>
    Public Property Connection As MySqlConnection
    ''' <summary>
    ''' Gets or sets an optional local transaction. When specified, its connection is used automatically unless <see cref="Connection"/> is also supplied.
    ''' </summary>
    Public Property Transaction As MySqlTransaction
    ''' <summary>
    ''' Gets or sets an optional command timeout, in seconds. A null value uses the provider default.
    ''' </summary>
    Public Property CommandTimeout As Integer?
End Class
