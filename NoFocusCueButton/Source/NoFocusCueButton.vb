Imports System.ComponentModel
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
    ''' Gets or sets the tooltip text displayed when hovering over this button.
    ''' </summary>
    ''' <value>
    ''' A <see cref="String"/> containing the tooltip text, or <see langword="Nothing"/> if no tooltip is set.
    ''' </value>
    <Category("NoFocusCueButton")>
    <Description("Specifies the text displayed in the tooltip when the mouse hovers over this button.")>
    <DefaultValue("")>
    Public Overridable Property TooltipText As String
        Get
            Return Tooltip.GetToolTip(Me)
        End Get
        Set(value As String)
            Tooltip.SetToolTip(Me, value)
        End Set
    End Property
    Private ReadOnly Tooltip As New ToolTip With {.ShowAlways = True}
End Class
