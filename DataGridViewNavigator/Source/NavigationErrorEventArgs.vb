''' <summary>Error information for a navigation initiated by a ToolStrip button.</summary>
Public NotInheritable Class NavigationErrorEventArgs
    Inherits EventArgs

    Public Sub New(Direction As NavigationDirection, [Error] As Exception)
        If [Error] Is Nothing Then Throw New ArgumentNullException(NameOf([Error]))
        Me.Direction = Direction
        Me.[Error] = [Error]
    End Sub

    ''' <summary>The attempted navigation direction.</summary>
    Public ReadOnly Property Direction As NavigationDirection

    ''' <summary>The exception raised by a navigation callback or operation.</summary>
    Public ReadOnly Property [Error] As Exception
End Class
