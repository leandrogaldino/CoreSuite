''' <summary>
''' Provides restore progress, error handling, parallel processing, and timeout settings.
''' </summary>
Public NotInheritable Class MySqlRestoreOptions
    ''' <summary>
    ''' Gets or sets the interval used by MySqlBackup.NET to report progress. The default is 1.
    ''' </summary>
    Public Property ProgressReportInterval As Integer = 1
    ''' <summary>
    ''' Gets or sets an optional progress receiver that receives values from 0 through 100.
    ''' </summary>
    Public Property Progress As IProgress(Of Integer)
    ''' <summary>
    ''' Gets or sets an optional command timeout, in seconds. A null value uses the provider default.
    ''' </summary>
    Public Property CommandTimeout As Integer?
    ''' <summary>
    ''' Gets or sets a value indicating whether SQL execution errors are ignored during restoration. The default is False.
    ''' </summary>
    Public Property IgnoreSqlError As Boolean = False
    ''' <summary>
    ''' Gets or sets a value indicating whether SQL statements are processed in parallel during restoration. The default is False.
    ''' </summary>
    Public Property EnableParallelProcessing As Boolean = True
    ''' <summary>
    ''' Gets or sets the optional file path used by MySqlBackup.NET to log SQL errors encountered during restoration.
    ''' </summary>
    Public Property ErrorLogFile As String = String.Empty
End Class
