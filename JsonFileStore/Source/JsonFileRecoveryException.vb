Imports System.IO
''' <summary>
''' Represents an error that occurs when a JSON file cannot be loaded and its backup cannot complete the recovery operation.
''' </summary>
Public NotInheritable Class JsonFileRecoveryException
    Inherits IOException
    Private ReadOnly _PrimaryException As Exception
    Private ReadOnly _RecoveryException As Exception
    ''' <summary>
    ''' Gets the exception raised while loading the primary JSON file.
    ''' </summary>
    ''' <value>The primary-file exception.</value>
    Public ReadOnly Property PrimaryException As Exception
        Get
            Return _PrimaryException
        End Get
    End Property
    ''' <summary>
    ''' Gets the exception raised while loading the backup or restoring the primary file.
    ''' </summary>
    ''' <value>The backup or restoration exception.</value>
    Public ReadOnly Property RecoveryException As Exception
        Get
            Return _RecoveryException
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="JsonFileRecoveryException"/> class.
    ''' </summary>
    ''' <param name="filePath">The path of the primary JSON file.</param>
    ''' <param name="backupPath">The path of the backup file.</param>
    ''' <param name="primaryException">The exception raised while loading the primary file.</param>
    ''' <param name="recoveryException">The exception raised while attempting recovery.</param>
    Friend Sub New(filePath As String, backupPath As String, primaryException As Exception, recoveryException As Exception)
        MyBase.New($"The JSON file '{filePath}' could not be loaded or recovered from '{backupPath}'.", primaryException)
        _PrimaryException = primaryException
        _RecoveryException = recoveryException
    End Sub
End Class
