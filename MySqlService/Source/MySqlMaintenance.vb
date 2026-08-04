Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Threading
Imports MySql.Data.MySqlClient
''' <summary>
''' Creates the configured database and performs managed SQL backup and restore operations.
''' </summary>
Public NotInheritable Class MySqlMaintenance
    Private ReadOnly _Client As MySqlClient
    ''' <summary>
    ''' Occurs when backup progress changes. Asynchronous operations raise this event from a worker thread.
    ''' </summary>
    Public Event BackupProgressChanged As EventHandler(Of ProgressChangedEventArgs)
    ''' <summary>
    ''' Occurs when restore progress changes. Asynchronous operations raise this event from a worker thread.
    ''' </summary>
    Public Event RestoreProgressChanged As EventHandler(Of ProgressChangedEventArgs)
    Friend Sub New(Client As MySqlClient)
        ArgumentNullException.ThrowIfNull(Client)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Creates the configured database using a server-level connection.
    ''' </summary>
    ''' <param name="Options">Optional character set, collation, existence, and timeout settings.</param>
    Public Sub ExecuteCreateDatabase(Optional Options As MySqlCreateDatabaseOptions = Nothing)
        Dim ActualOptions As MySqlCreateDatabaseOptions = ValidateCreateDatabaseOptions(Options)
        Using Connection As MySqlConnection = _Client.CreateServerConnection()
            Connection.Open()
            Using Command As MySqlCommand = CreateDatabaseCommand(Connection, ActualOptions)
                Command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    ''' <summary>
    ''' Asynchronously creates the configured database using a server-level connection.
    ''' </summary>
    ''' <param name="Options">Optional character set, collation, existence, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening or command execution.</param>
    ''' <returns>A task representing the database creation operation.</returns>
    Public Async Function ExecuteCreateDatabaseAsync(Optional Options As MySqlCreateDatabaseOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        Dim ActualOptions As MySqlCreateDatabaseOptions = ValidateCreateDatabaseOptions(Options)
        Using Connection As MySqlConnection = _Client.CreateServerConnection()
            Await Connection.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Command As MySqlCommand = CreateDatabaseCommand(Connection, ActualOptions)
                Await Command.ExecuteNonQueryAsync(CancellationToken).ConfigureAwait(False)
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Exports the configured database to an atomically replaced SQL file.
    ''' </summary>
    ''' <param name="FilePath">The destination SQL file path.</param>
    ''' <param name="Options">Optional export content, progress, overwrite, and timeout settings.</param>
    Public Sub ExecuteBackup(FilePath As String, Optional Options As MySqlBackupOptions = Nothing)
        ExecuteBackupCore(FilePath, ValidateBackupOptions(Options), CancellationToken.None)
    End Sub
    ''' <summary>
    ''' Asynchronously exports the configured database to an atomically replaced SQL file without blocking the calling thread.
    ''' </summary>
    ''' <param name="FilePath">The destination SQL file path.</param>
    ''' <param name="Options">Optional export content, progress, overwrite, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel the export.</param>
    ''' <returns>A task representing the backup operation.</returns>
    Public Function ExecuteBackupAsync(FilePath As String, Optional Options As MySqlBackupOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        Dim ActualOptions As MySqlBackupOptions = ValidateBackupOptions(Options)
        Return Task.Run(Sub() ExecuteBackupCore(FilePath, ActualOptions, CancellationToken), CancellationToken)
    End Function
    ''' <summary>
    ''' Imports a SQL backup file into the configured database.
    ''' </summary>
    ''' <param name="FilePath">The source SQL file path.</param>
    ''' <param name="Options">Optional progress and timeout settings.</param>
    Public Sub ExecuteRestore(FilePath As String, Optional Options As MySqlRestoreOptions = Nothing)
        ExecuteRestoreCore(FilePath, ValidateRestoreOptions(Options), CancellationToken.None)
    End Sub
    ''' <summary>
    ''' Asynchronously imports a SQL backup file without blocking the calling thread.
    ''' </summary>
    ''' <param name="FilePath">The source SQL file path.</param>
    ''' <param name="Options">Optional progress and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel the import.</param>
    ''' <returns>A task representing the restore operation.</returns>
    Public Function ExecuteRestoreAsync(FilePath As String, Optional Options As MySqlRestoreOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        Dim ActualOptions As MySqlRestoreOptions = ValidateRestoreOptions(Options)
        Return Task.Run(Sub() ExecuteRestoreCore(FilePath, ActualOptions, CancellationToken), CancellationToken)
    End Function
    Private Function CreateDatabaseCommand(Connection As MySqlConnection, Options As MySqlCreateDatabaseOptions) As MySqlCommand
        Dim CharacterSet As String = ValidateSqlToken(Options.CharacterSet, NameOf(Options.CharacterSet))
        Dim Collation As String = ValidateSqlToken(Options.Collation, NameOf(Options.Collation))
        Dim IfNotExistsClause As String = If(Options.IfNotExists, "IF NOT EXISTS ", String.Empty)
        Dim Sql As String = $"CREATE DATABASE {IfNotExistsClause}{MySqlSql.QuoteSingleIdentifier(_Client.Database, "database")} CHARACTER SET {CharacterSet} COLLATE {Collation};"
        Dim Command As New MySqlCommand(Sql, Connection)
        If Options.CommandTimeout.HasValue Then Command.CommandTimeout = Options.CommandTimeout.Value
        Return Command
    End Function
    Private Sub ExecuteBackupCore(FilePath As String, Options As MySqlBackupOptions, CancellationToken As CancellationToken)
        Dim TargetPath As String = PrepareBackupPath(FilePath, Options)
        Dim TemporaryPath As String = CreateTemporaryPath(TargetPath)
        Try
            CancellationToken.ThrowIfCancellationRequested()
            Using Connection As MySqlConnection = _Client.CreateDatabaseConnection()
                Connection.OpenAsync(CancellationToken).GetAwaiter().GetResult()
                Using Command As New MySqlCommand With {.Connection = Connection}
                    If Options.CommandTimeout.HasValue Then Command.CommandTimeout = Options.CommandTimeout.Value
                    Using Backup As New MySqlBackup(Command)
                        Backup.ExportInfo.IntervalForProgressReport = Options.ProgressReportInterval
                        Backup.ExportInfo.ExportProcedures = Options.ExportProcedures
                        Backup.ExportInfo.ExportFunctions = Options.ExportFunctions
                        Backup.ExportInfo.ExportTriggers = Options.ExportTriggers
                        Dim LastProgress As Integer = ReportBackupProgress(0, Options.Progress, -1)
                        AddHandler Backup.ExportProgressChanged, Sub(sender, eventArgs) LastProgress = ReportBackupProgress(ClampPercentage(eventArgs.CurrentRowIndexInAllTables, eventArgs.TotalRowsInAllTables), Options.Progress, LastProgress)
                        Using cancellationRegistration As CancellationTokenRegistration = CancellationToken.Register(Sub() CancelOperation(Backup, Connection))
                            Backup.ExportToFile(TemporaryPath)
                        End Using
                        CancellationToken.ThrowIfCancellationRequested()
                        LastProgress = ReportBackupProgress(100, Options.Progress, LastProgress)
                    End Using
                End Using
            End Using
            MoveBackupIntoPlace(TemporaryPath, TargetPath, Options.Overwrite)
        Catch ex As Exception When CancellationToken.IsCancellationRequested
            Throw New OperationCanceledException("The database backup was canceled.", ex, CancellationToken)
        Finally
            TryDeleteFile(TemporaryPath)
        End Try
    End Sub
    Private Sub ExecuteRestoreCore(FilePath As String, Options As MySqlRestoreOptions, CancellationToken As CancellationToken)
        Dim FullPath As String = ValidateRestorePath(FilePath)
        Dim TotalBytes As Long = New FileInfo(FullPath).Length
        Try
            CancellationToken.ThrowIfCancellationRequested()
            Using Connection As MySqlConnection = _Client.CreateDatabaseConnection()
                Connection.OpenAsync(CancellationToken).GetAwaiter().GetResult()
                Using Command As New MySqlCommand With {.Connection = Connection}
                    If Options.CommandTimeout.HasValue Then Command.CommandTimeout = Options.CommandTimeout.Value
                    Using Backup As New MySqlBackup(Command)
                        Backup.ImportInfo.IntervalForProgressReport = Options.ProgressReportInterval
                        Dim LastProgress As Integer = ReportRestoreProgress(0, Options.Progress, -1)
                        AddHandler Backup.ImportProgressChanged, Sub(sender, eventArgs) LastProgress = ReportRestoreProgress(ClampPercentage(eventArgs.CurrentBytes, TotalBytes), Options.Progress, LastProgress)
                        Using CancellationRegistration As CancellationTokenRegistration = CancellationToken.Register(Sub() CancelOperation(Backup, Connection))
                            Backup.ImportFromFile(FullPath)
                        End Using
                        CancellationToken.ThrowIfCancellationRequested()
                        LastProgress = ReportRestoreProgress(100, Options.Progress, LastProgress)
                    End Using
                End Using
            End Using
        Catch ex As Exception When CancellationToken.IsCancellationRequested
            Throw New OperationCanceledException("The database restore was canceled.", ex, CancellationToken)
        End Try
    End Sub
    Private Shared Function ValidateCreateDatabaseOptions(Options As MySqlCreateDatabaseOptions) As MySqlCreateDatabaseOptions
        Dim ActualOptions As MySqlCreateDatabaseOptions = If(Options, New MySqlCreateDatabaseOptions())
        ValidateCommandTimeout(ActualOptions.CommandTimeout, NameOf(Options))
        ValidateSqlToken(ActualOptions.CharacterSet, NameOf(ActualOptions.CharacterSet))
        ValidateSqlToken(ActualOptions.Collation, NameOf(ActualOptions.Collation))
        Return ActualOptions
    End Function
    Private Shared Function ValidateBackupOptions(options As MySqlBackupOptions) As MySqlBackupOptions
        Dim ActualOptions As MySqlBackupOptions = If(options, New MySqlBackupOptions())
        RequirePositive(ActualOptions.ProgressReportInterval, NameOf(ActualOptions.ProgressReportInterval))
        ValidateCommandTimeout(ActualOptions.CommandTimeout, NameOf(options))
        Return ActualOptions
    End Function
    Private Shared Function ValidateRestoreOptions(options As MySqlRestoreOptions) As MySqlRestoreOptions
        Dim ActualOptions As MySqlRestoreOptions = If(options, New MySqlRestoreOptions())
        RequirePositive(ActualOptions.ProgressReportInterval, NameOf(ActualOptions.ProgressReportInterval))
        ValidateCommandTimeout(ActualOptions.CommandTimeout, NameOf(options))
        Return ActualOptions
    End Function
    Private Shared Sub ValidateCommandTimeout(CommandTimeout As Integer?, ParameterName As String)
        If CommandTimeout.HasValue AndAlso CommandTimeout.Value < 0 Then Throw New ArgumentOutOfRangeException(ParameterName, CommandTimeout.Value, "CommandTimeout cannot be negative.")
    End Sub
    Private Shared Function PrepareBackupPath(FilePath As String, Options As MySqlBackupOptions) As String
        Dim FullPath As String = Path.GetFullPath(RequireValue(FilePath, NameOf(FilePath)))
        If String.IsNullOrWhiteSpace(Path.GetFileName(FullPath)) Then Throw New ArgumentException("The backup path must include a file name.", NameOf(FilePath))
        Dim DirectoryPath As String = Path.GetDirectoryName(FullPath)
        If String.IsNullOrWhiteSpace(DirectoryPath) Then DirectoryPath = Directory.GetCurrentDirectory()
        Directory.CreateDirectory(DirectoryPath)
        If File.Exists(FullPath) AndAlso Not Options.Overwrite Then Throw New IOException($"The backup file already exists: {FullPath}")
        Return FullPath
    End Function
    Private Shared Function ValidateRestorePath(FilePath As String) As String
        Dim FullPath As String = Path.GetFullPath(RequireValue(FilePath, NameOf(FilePath)))
        If Not File.Exists(FullPath) Then Throw New FileNotFoundException("The backup file was not found.", FullPath)
        If New FileInfo(FullPath).Length = 0 Then Throw New InvalidDataException("The backup file is empty.")
        Return FullPath
    End Function
    Private Shared Function CreateTemporaryPath(TargetPath As String) As String
        Dim DirectoryPath As String = Path.GetDirectoryName(TargetPath)
        Return Path.Combine(DirectoryPath, $".{Path.GetFileName(TargetPath)}.{Guid.NewGuid():N}.tmp")
    End Function
    Private Shared Sub MoveBackupIntoPlace(SourcePath As String, TargetPath As String, Overwrite As Boolean)
        File.Move(SourcePath, TargetPath, Overwrite)
    End Sub
    Private Shared Sub CancelOperation(Backup As MySqlBackup, Connection As MySqlConnection)
        Try
            Backup.StopAllProcess()
        Catch ex As Exception
        End Try
        SafeClose(Connection)
    End Sub
    Private Shared Sub SafeClose(Connection As MySqlConnection)
        Try
            If Connection.State <> ConnectionState.Closed Then Connection.Close()
        Catch ex As MySqlException
        Catch ex As InvalidOperationException
        End Try
    End Sub
    Private Shared Sub TryDeleteFile(FilePath As String)
        Try
            If File.Exists(FilePath) Then File.Delete(FilePath)
        Catch ex As IOException
        Catch ex As UnauthorizedAccessException
        End Try
    End Sub
    Private Function ReportBackupProgress(Percentage As Integer, Progress As IProgress(Of Integer), LastProgress As Integer) As Integer
        If Percentage = LastProgress Then Return LastProgress
        RaiseEvent BackupProgressChanged(Me, New ProgressChangedEventArgs(Percentage, Nothing))
        Progress?.Report(Percentage)
        Return Percentage
    End Function
    Private Function ReportRestoreProgress(Percentage As Integer, Progress As IProgress(Of Integer), LastProgress As Integer) As Integer
        If Percentage = LastProgress Then Return LastProgress
        RaiseEvent RestoreProgressChanged(Me, New ProgressChangedEventArgs(Percentage, Nothing))
        Progress?.Report(Percentage)
        Return Percentage
    End Function
End Class
