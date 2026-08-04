''' <summary>
''' Provides restore progress and timeout settings.
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
End Class
