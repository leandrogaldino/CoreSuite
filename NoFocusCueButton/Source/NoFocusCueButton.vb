
Imports System.ComponentModel
Imports System.Windows.Forms

''' <summary>
''' Represents a button that suppresses focus cues and provides
''' built-in tooltip support.
''' </summary>
Public Class NoFocusCueButton
    Inherits Button

    Private ReadOnly _ToolTipComponent As New ToolTip With {
        .ShowAlways = True
    }

    Private _ToolTipTextValue As String = String.Empty

    ''' <summary>
    ''' Gets a value indicating that focus cues should never
    ''' be displayed for this button.
    ''' </summary>
    Protected Overrides ReadOnly Property ShowFocusCues As Boolean
        Get
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the tooltip text displayed when the
    ''' mouse pointer hovers over this button.
    ''' </summary>
    <Category("Behavior")>
    <Description("Specifies the tooltip text displayed when hovering over the button.")>
    <DefaultValue("")>
    <Localizable(True)>
    Public Property TooltipText As String
        Get
            Return _ToolTipTextValue
        End Get
        Set(Value As String)
            Dim NewValue As String = If(Value, String.Empty)

            If String.Equals(_ToolTipTextValue, NewValue, StringComparison.Ordinal) Then
                Return
            End If

            _ToolTipTextValue = NewValue
            _ToolTipComponent.SetToolTip(Me, NewValue)
        End Set
    End Property

    ''' <summary>
    ''' Releases the resources used by the button.
    ''' </summary>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            _ToolTipComponent.Dispose()
        End If

        MyBase.Dispose(Disposing)
    End Sub

End Class
