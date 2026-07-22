
Friend Class EditorServiceEventArgs
    Inherits EventArgs

    Public Sub New(ByVal colorUI As Control)
        _ColorUI = colorUI
    End Sub

    Private _ColorUI As Control

    Public ReadOnly Property ColorUI As Control
        Get
            Return _ColorUI
        End Get
    End Property
End Class