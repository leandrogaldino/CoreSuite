''' <summary>
''' Provides a custom renderer for <see cref="ToolStrip"/> controls.
''' </summary>
''' <remarks>
''' This renderer allows customization of the visual appearance of ToolStrip items,
''' including the option to display the default checked state background for checked
''' buttons.
''' </remarks>
Public Class CToolStripRender
    Inherits ToolStripSystemRenderer
    ''' <summary>
    ''' Gets or sets a value indicating whether the default checked button background
    ''' should be displayed for checked <see cref="ToolStripButton"/> controls.
    ''' </summary>
    ''' <remarks>
    ''' When enabled, checked buttons are rendered using the base ToolStrip renderer,
    ''' allowing the default selection rectangle or highlight to be displayed.
    ''' </remarks>
    Public Property ShowButtonCheckedRectangle As Boolean
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CToolStripRender"/> class.
    ''' </summary>
    ''' <remarks>
    ''' Creates a renderer with the default ToolStrip rendering behavior.
    ''' </remarks>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Prevents the default ToolStrip border from being rendered.
    ''' </summary>
    ''' <remarks>
    ''' This override removes the standard ToolStrip border rendering, allowing the
    ''' control appearance to be fully customized.
    ''' </remarks>
    ''' <param name="e">
    ''' Contains information about the ToolStrip rendering operation.
    ''' </param>
    <DebuggerStepThrough>
    Protected Overrides Sub OnRenderToolStripBorder(ByVal e As ToolStripRenderEventArgs)
    End Sub
    ''' <summary>
    ''' Renders the background of a ToolStrip item according to its checked state.
    ''' </summary>
    ''' <remarks>
    ''' Checked ToolStrip buttons only display their default background rendering when
    ''' <see cref="ShowButtonCheckedRectangle"/> is enabled. Other ToolStrip items use
    ''' the default renderer behavior.
    ''' </remarks>
    ''' <param name="e">
    ''' Contains information about the ToolStrip item rendering operation.
    ''' </param>
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