Imports System.ComponentModel
Imports System.Data.Common
Imports System.IO
Imports System.Text.RegularExpressions
Imports MySql.Data.MySqlClient

Public Class MySqlMaintenance

    Public Event BackupProgressChanged As EventHandler(Of ProgressChangedEventArgs)

    Public Event RestoreProgressChanged As EventHandler(Of ProgressChangedEventArgs)

    Private ReadOnly _Client As MySqlClient

    Friend Sub New(Client As MySqlClient)
        _Client = Client
    End Sub

    ''' <summary>
    ''' Creates a database on the MySQL server according to the data provided to <see cref="MySqlClient"/>.
    ''' </summary>
    ''' <param name="IfNotExists">If true, adds IF NOT EXISTS to avoid errors if the database already exists.</param>
    ''' <param name="Charset">Database charset (default: utf8mb4).</param>
    ''' <param name="Collation">Database collation (default: utf8mb4_unicode_ci).</param>
    ''' <param name="Connection">Optional connection. If not provided, the class will create one.</param>
    ''' <returns>A task representing the asynchronous operation.</returns>
    Public Async Function ExecuteCreateDatabaseAsync(Optional IfNotExists As Boolean = True, Optional Charset As String = "utf8mb4", Optional Collation As String = "utf8mb4_unicode_ci", Optional Connection As DbConnection = Nothing) As Task
        If String.IsNullOrWhiteSpace(_Client.Database) Then
            Throw New ArgumentException("DatabaseName inválido.", NameOf(_Client.Database))
        End If
        If Not Regex.IsMatch(_Client.Database, "^[a-zA-Z0-9_]+$") Then
            Throw New ArgumentException("DatabaseName contém caracteres inválidos.")
        End If
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateServerConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Dim Query As String = $"CREATE DATABASE {(If(IfNotExists, "IF NOT EXISTS ", ""))}`{_Client.Database}` CHARACTER SET {Charset} COLLATE {Collation};"
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = Query
                Await Command.ExecuteNonQueryAsync()
            End Using
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Restores the database using only the backup file path.
    ''' </summary>
    ''' <param name="FilePath">Path to the backup file to be restored.</param>
    ''' <returns>An asynchronous task.</returns>
    Public Overloads Async Function ExecuteRestoreAsync(FilePath As String) As Task
        Await ExecuteRestoreAsync(FilePath, Nothing, Nothing)
    End Function

    ''' <summary>
    ''' Restores the database using a file path and an existing connection.
    ''' </summary>
    ''' <param name="FilePath">Path to the backup file to be restored.</param>
    ''' <param name="Connection">Existing connection.</param>
    ''' <returns>An asynchronous task.</returns>
    Public Overloads Async Function ExecuteRestoreAsync(FilePath As String, Connection As DbConnection) As Task
        Await ExecuteRestoreAsync(FilePath, Connection, Nothing)
    End Function

    ''' <summary>
    ''' Restores the database using a file path and progress reporting.
    ''' </summary>
    ''' <param name="FilePath">Path to the backup file to be restored.</param>
    ''' <param name="Progress">Optional progress reported during restore (0 to 100).</param>
    ''' <returns>An asynchronous task.</returns>
    Public Overloads Async Function ExecuteRestoreAsync(FilePath As String, Progress As IProgress(Of Integer)) As Task
        Await ExecuteRestoreAsync(FilePath, Nothing, Progress)
    End Function

    Private Overloads Async Function ExecuteRestoreAsync(FilePath As String, Optional Connection As DbConnection = Nothing, Optional Progress As IProgress(Of Integer) = Nothing) As Task
        If String.IsNullOrWhiteSpace(FilePath) Then
            Throw New ArgumentException("FilePath inválido.", NameOf(FilePath))
        End If
        If Not File.Exists(FilePath) Then
            Throw New FileNotFoundException("Arquivo de backup não encontrado.", FilePath)
        End If

        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If

        Try
            Using Cmd As New MySqlCommand(Nothing, Connection)
                Using Bkp As New MySqlBackup(Cmd)
                    Bkp.ImportInfo.IntervalForProgressReport = 1

                    AddHandler Bkp.ImportProgressChanged, Sub(sender, e)
                                                              Dim TotalBytes As Long = New FileInfo(FilePath).Length
                                                              Dim CurrentBytes As Long = e.CurrentBytes
                                                              Dim Percent As Integer = CInt((CurrentBytes / TotalBytes) * 100)
                                                              RaiseEvent RestoreProgressChanged(Me, New ProgressChangedEventArgs(Percent, Nothing))
                                                              Progress?.Report(Percent)
                                                          End Sub

                    AddHandler Bkp.ImportCompleted, Sub(sender, e)
                                                        RaiseEvent RestoreProgressChanged(Me, New ProgressChangedEventArgs(100, Nothing))
                                                        Progress?.Report(100)
                                                    End Sub

                    Bkp.ImportFromFile(FilePath)
                End Using
            End Using
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Performs a backup using only the file path.
    ''' </summary>
    ''' <param name="FilePath">Path where the backup file will be saved.</param>
    ''' <returns>An asynchronous task.</returns>
    Public Overloads Async Function ExecuteBackupAsync(FilePath As String) As Task
        Await ExecuteBackupAsync(FilePath, Nothing, Nothing)
    End Function

    ''' <summary>
    ''' Backup usando caminho do arquivo e conexão.
    ''' </summary>
    ''' <param name="FilePath">Caminho do arquivo de backup.</param>
    ''' <param name="Connection">Conexão existente.</param>
    ''' <returns>Task assíncrona.</returns>
    Public Overloads Async Function ExecuteBackupAsync(FilePath As String, Connection As DbConnection) As Task
        Await ExecuteBackupAsync(FilePath, Connection, Nothing)
    End Function

    ''' <summary>
    ''' Performs a backup using a file path and progress reporting.
    ''' </summary>
    ''' <param name="FilePath">Path where the backup file will be saved.</param>
    ''' <param name="Progress">Optional progress reporting.</param>
    ''' <returns>An asynchronous task.</returns>
    Public Overloads Async Function ExecuteBackupAsync(FilePath As String, Progress As IProgress(Of Integer)) As Task
        Await ExecuteBackupAsync(FilePath, Nothing, Progress)
    End Function

    Private Overloads Async Function ExecuteBackupAsync(FilePath As String, Optional Connection As DbConnection = Nothing, Optional Progress As IProgress(Of Integer) = Nothing) As Task
        If String.IsNullOrWhiteSpace(FilePath) Then
            Throw New ArgumentException("FilePath inválido.", NameOf(FilePath))
        End If

        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Using Cmd As New MySqlCommand(Nothing, Connection)
                Using Bkp As New MySqlBackup(Cmd)
                    Bkp.ExportInfo.IntervalForProgressReport = 1

                    AddHandler Bkp.ExportProgressChanged, Sub(sender, e)
                                                              Dim TotalRows As Long = e.TotalRowsInAllTables
                                                              Dim CurrentRow As Long = e.CurrentRowIndexInAllTables
                                                              Dim Percent As Integer = CInt((CurrentRow / TotalRows) * 100)
                                                              RaiseEvent BackupProgressChanged(Me, New ProgressChangedEventArgs(Percent, Nothing))
                                                              Progress?.Report(Percent)
                                                          End Sub

                    AddHandler Bkp.ExportCompleted, Sub(sender, e)
                                                        RaiseEvent BackupProgressChanged(Me, New ProgressChangedEventArgs(100, Nothing))
                                                        Progress?.Report(100)
                                                    End Sub

                    Using Stream As New FileStream(FilePath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, True)
                        Bkp.ExportToStream(Stream)
                        Await Stream.FlushAsync()
                    End Using
                End Using
            End Using
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

End Class