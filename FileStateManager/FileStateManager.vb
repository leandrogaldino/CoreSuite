Imports System.IO
Imports ChinhDo.Transactions
Public Class FileStateManager
    Private ReadOnly _TargetDirectory As String
    Private _OriginalFile As String
    Private _CurrentFile As String
    Public Shared Property TempDirectory As String
    Public ReadOnly Property TargetDirectory As String
        Get
            Return _TargetDirectory
        End Get
    End Property

    Public ReadOnly Property OriginalFile As String
        Get
            Return _OriginalFile
        End Get
    End Property
    Public ReadOnly Property CurrentFile As String
        Get
            Return _CurrentFile
        End Get
    End Property
    Public Sub New(TargetDirectory As String)
        TempDirectory = TempDirectory
        _TargetDirectory = TargetDirectory
    End Sub
    Public Sub SetCurrentFile(Filename As String, Optional AsOriginal As Boolean = False)
        If String.IsNullOrEmpty(Filename) Or String.IsNullOrWhiteSpace(Filename) Then
            Filename = Nothing
        End If
        _CurrentFile = Filename
        If AsOriginal Then
            _OriginalFile = Filename
        End If
    End Sub
    Public Sub Execute()
        Dim FileManager As TxFileManager
        If String.IsNullOrEmpty(TempDirectory) Then
            FileManager = New TxFileManager()
        Else
            FileManager = New TxFileManager(TempDirectory)
        End If
        If String.IsNullOrEmpty(CurrentFile) And Not String.IsNullOrEmpty(_OriginalFile) Then
            'se nao tem atual mas tem original
            'deleta o original
            FileManager.Delete(_OriginalFile)
            'seta original como nada
            _OriginalFile = Nothing
        ElseIf Not String.IsNullOrEmpty(CurrentFile) And String.IsNullOrEmpty(_OriginalFile) Then
            'se tem atual mas nao tem original
            'copia o atual
            FileManager.Copy(CurrentFile, Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile)), False)
            'seta original igual atual
            _CurrentFile = Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile))
            _OriginalFile = CurrentFile
        ElseIf Not String.IsNullOrEmpty(CurrentFile) And Not String.IsNullOrEmpty(_OriginalFile) Then
            'se tem atual e tem original
            If CurrentFile <> _OriginalFile Then
                'se o atual e diferente do original
                'deleta o original
                FileManager.Delete(_OriginalFile)
                'copia o atual
                FileManager.Copy(CurrentFile, Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile)), False)
                'seta original igual atual
                _CurrentFile = Path.Combine(_TargetDirectory, Path.GetFileName(CurrentFile))
                _OriginalFile = CurrentFile
            End If
        End If
    End Sub

    Public Function Clone() As FileStateManager
        Dim Fsm As New FileStateManager(Me._TargetDirectory)
        Fsm.SetCurrentFile(Me._CurrentFile, False)
        Fsm.SetCurrentFile(Me._OriginalFile, True)
        Return Fsm
    End Function
End Class
