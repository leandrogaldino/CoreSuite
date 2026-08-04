Imports System.IO
Imports ChinhDo.Transactions
''' <summary>
''' Manages the transactional state of a file within a target directory, tracking
''' the original and current file paths and applying changes (copy or delete)
''' using a transactional file manager.
''' </summary>
Public Class FileStateManager
    Private ReadOnly _TargetDirectory As String
    Private _OriginalFile As String
    Private _CurrentFile As String
    ''' <summary>
    ''' Gets or sets the temporary directory used by the transactional file manager.
    ''' </summary>
    Public Shared Property TempDirectory As String
    ''' <summary>
    ''' Gets the target directory where the current file will be stored when changes are applied.
    ''' </summary>
    Public ReadOnly Property TargetDirectory As String
        Get
            Return _TargetDirectory
        End Get
    End Property
    ''' <summary>
    ''' Gets the original file path before any pending changes are applied.
    ''' </summary>
    Public ReadOnly Property OriginalFile As String
        Get
            Return _OriginalFile
        End Get
    End Property
    ''' <summary>
    ''' Gets the current file path representing the pending state.
    ''' </summary>
    Public ReadOnly Property CurrentFile As String
        Get
            Return _CurrentFile
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FileStateManager"/> class
    ''' for the specified target directory.
    ''' </summary>
    ''' <param name="TargetDirectory">The directory where the current file will be stored when changes are applied.</param>
    Public Sub New(TargetDirectory As String)
        TempDirectory = TempDirectory
        _TargetDirectory = TargetDirectory
    End Sub
    ''' <summary>
    ''' Sets the current file path, optionally marking it as the original file as well.
    ''' </summary>
    ''' <param name="Filename">The file path to set as current. If null, empty, or whitespace, it is treated as <c>Nothing</c>.</param>
    ''' <param name="AsOriginal">If <c>True</c>, also sets the specified file as the original file.</param>
    Public Sub SetCurrentFile(Filename As String, Optional AsOriginal As Boolean = False)
        If String.IsNullOrEmpty(Filename) Or String.IsNullOrWhiteSpace(Filename) Then
            Filename = Nothing
        End If
        _CurrentFile = Filename
        If AsOriginal Then
            _OriginalFile = Filename
        End If
    End Sub
    ''' <summary>
    ''' Applies the pending file state changes transactionally, copying the current file
    ''' to the target directory and/or deleting the original file as needed.
    ''' </summary>
    Public Sub Execute()
        Dim FileManager As TxFileManager
        If String.IsNullOrEmpty(TempDirectory) Then
            FileManager = New TxFileManager()
        Else
            FileManager = New TxFileManager(TempDirectory)
        End If
        If String.IsNullOrEmpty(CurrentFile) And Not String.IsNullOrEmpty(_OriginalFile) Then
            'If there is no current file but there is an original, delete the original.
            FileManager.Delete(_OriginalFile)
            'Set original to nothing.
            _OriginalFile = Nothing
        ElseIf Not String.IsNullOrEmpty(CurrentFile) And String.IsNullOrEmpty(_OriginalFile) Then
            'If there is a current file but no original, copy the current file.
            FileManager.Copy(CurrentFile, Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile)), False)
            'Set original equal to current.
            _CurrentFile = Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile))
            _OriginalFile = CurrentFile
        ElseIf Not String.IsNullOrEmpty(CurrentFile) And Not String.IsNullOrEmpty(_OriginalFile) Then
            'If there is both a current file and an original
            If CurrentFile <> _OriginalFile Then
                'If the current file is different from the original, delete the original.
                FileManager.Delete(_OriginalFile)
                'Copy the current file.
                FileManager.Copy(CurrentFile, Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile)), False)
                'Set original equal to current.
                _CurrentFile = Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile))
                _OriginalFile = CurrentFile
            End If
        End If
    End Sub
    ''' <summary>
    ''' Creates a copy of the current <see cref="FileStateManager"/> instance, preserving
    ''' its target directory, current file, and original file.
    ''' </summary>
    ''' <returns>A new <see cref="FileStateManager"/> instance with the same state.</returns>
    Public Function Clone() As FileStateManager
        Dim Fsm As New FileStateManager(Me._TargetDirectory)
        Fsm.SetCurrentFile(Me._CurrentFile, False)
        Fsm.SetCurrentFile(Me._OriginalFile, True)
        Return Fsm
    End Function
End Class
