Imports System.Windows.Forms

Public Class NoFocusCueButton
    Inherits Button
    Private ReadOnly Tooltip As New ToolTip

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

End Class