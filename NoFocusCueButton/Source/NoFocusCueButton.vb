''' <summary>
''' Represents a <see cref="Button"/> that suppresses the default focus rectangle
''' and supports an associated tooltip.
''' </summary>
Public Class NoFocusCueButton
    Inherits Button
    ''' <summary>
    ''' Gets a value indicating that focus cues should never be shown for this button.
    ''' </summary>
    Protected Overrides ReadOnly Property ShowFocusCues As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for this button.
    ''' </summary>
    Public Property TooltipText As String
        Get
            Return Tooltip.GetToolTip(Me)
        End Get
        Set(value As String)
            Tooltip.SetToolTip(Me, value)
        End Set
    End Property
    Private ReadOnly Tooltip As New ToolTip
End Class
