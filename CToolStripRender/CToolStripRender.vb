Imports System.Windows.Forms

''' <summary>
''' Custom renderer for ToolStrip controls.
''' Allows customization of how ToolStrip buttons are drawn,
''' including optionally showing a rectangle around checked buttons.
''' </summary>
Public Class CToolStripRender
    Inherits ToolStripSystemRenderer

    ''' <summary>
    ''' Gets or sets a value indicating whether a rectangle should be drawn around checked buttons.
    ''' </summary>
    Public Property ShowButtonCheckedRectangle As Boolean

    ''' <summary>
    ''' Initializes a new instance of the <see cref="CToolStripRender"/> class.
    ''' </summary>
    Public Sub New()
    End Sub

    <DebuggerStepThrough>
    Protected Overrides Sub OnRenderToolStripBorder(ByVal e As ToolStripRenderEventArgs)
    End Sub

    <DebuggerStepThrough>
    Protected Overrides Sub OnRenderButtonBackground(e As ToolStripItemRenderEventArgs)
        Dim button = TryCast(e.Item, ToolStripButton)
        If (button IsNot Nothing AndAlso button.Checked) Then
            If ShowButtonCheckedRectangle Then
                MyBase.OnRenderButtonBackground(e)
            End If
        Else
            MyBase.OnRenderButtonBackground(e)
        End If
    End Sub

End Class