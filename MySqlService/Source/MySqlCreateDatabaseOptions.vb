''' <summary>
''' Provides database creation settings.
''' </summary>
Public NotInheritable Class MySqlCreateDatabaseOptions
    ''' <summary>
    ''' Gets or sets whether the command includes IF NOT EXISTS. The default is <see langword="True"/>.
    ''' </summary>
    Public Property IfNotExists As Boolean = True
    ''' <summary>
    ''' Gets or sets the database character set. The value is validated as a single SQL token. The default is utf8mb4.
    ''' </summary>
    Public Property CharacterSet As String = "utf8mb4"
    ''' <summary>
    ''' Gets or sets the database collation. The value is validated as a single SQL token. The default is utf8mb4_unicode_ci.
    ''' </summary>
    Public Property Collation As String = "utf8mb4_unicode_ci"
    ''' <summary>
    ''' Gets or sets an optional command timeout, in seconds. A null value uses the provider default.
    ''' </summary>
    Public Property CommandTimeout As Integer?
End Class
