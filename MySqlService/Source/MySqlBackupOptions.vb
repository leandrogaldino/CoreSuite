''' <summary>
''' Provides backup export, progress, overwrite, and timeout settings.
''' </summary>
Public NotInheritable Class MySqlBackupOptions
    ''' <summary>
    ''' Gets or sets whether an existing destination file may be replaced. The default is <see langword="True"/>.
    ''' </summary>
    Public Property Overwrite As Boolean = True
    ''' <summary>
    ''' Gets or sets whether stored procedures are exported. The default is <see langword="True"/>.
    ''' </summary>
    Public Property ExportProcedures As Boolean = True
    ''' <summary>
    ''' Gets or sets whether stored functions are exported. The default is <see langword="True"/>.
    ''' </summary>
    Public Property ExportFunctions As Boolean = True
    ''' <summary>
    ''' Gets or sets whether triggers are exported. The default is <see langword="True"/>.
    ''' </summary>
    Public Property ExportTriggers As Boolean = True
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
