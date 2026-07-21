Public Class NoFocusCueButton
    Inherits Button
    Protected Overrides ReadOnly Property ShowFocusCues As Boolean
        Get
            Return False
        End Get
    End Property
    Public Property TooltipText As String
        Get
            Return Tooltip.GetToolTip(Me)
        End Get
        Set(value As String)
            Tooltip.SetToolTip(Me, value)
        End Set
    End Property
    Private Tooltip As New ToolTip
End Class
