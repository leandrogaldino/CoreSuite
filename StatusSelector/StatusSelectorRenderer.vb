Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' Provides the professional renderer used by <see cref="StatusSelector"/>.
''' </summary>
''' <remarks>
''' The native <see cref="ToolStripProfessionalRenderer"/> remains responsible for menu item sizing and layout,
''' while this renderer supplies the configured hover colors, borders and selected-item appearance.
''' </remarks>
Friend NotInheritable Class StatusSelectorRenderer
    Inherits ToolStripProfessionalRenderer

    Private ReadOnly _Owner As StatusSelector

    ''' <summary>
    ''' Initializes a new renderer for the specified selector.
    ''' </summary>
    ''' <param name="owner">The selector whose appearance settings are used.</param>
    Public Sub New(owner As StatusSelector)
        MyBase.New(New StatusSelectorColorTable(owner))
        ArgumentNullException.ThrowIfNull(owner)
        _Owner = owner
        RoundedEdges = owner.MenuRoundedEdges
    End Sub

    ''' <summary>
    ''' Renders item text while preserving the native ToolStrip layout.
    ''' </summary>
    Protected Overrides Sub OnRenderItemText(e As ToolStripItemTextRenderEventArgs)
        If e.Item.Selected Then
            Dim StatusItem = TryCast(e.Item.Tag, StatusSelectorItem)
            If StatusItem IsNot Nothing AndAlso StatusItem.HoverForeColor <> Color.Empty Then
                e.TextColor = StatusItem.HoverForeColor
            ElseIf _Owner.HoverForeColor <> Color.Empty Then
                e.TextColor = _Owner.HoverForeColor
            End If
        End If
        MyBase.OnRenderItemText(e)
    End Sub

    ''' <summary>
    ''' Renders an item-specific hover background when one is configured.
    ''' </summary>
    Protected Overrides Sub OnRenderMenuItemBackground(e As ToolStripItemRenderEventArgs)
        If Not e.Item.Selected Then
            MyBase.OnRenderMenuItemBackground(e)
            Return
        End If

        Dim StatusItem = TryCast(e.Item.Tag, StatusSelectorItem)
        Dim HoverColor = _Owner.HoverBackColor
        If StatusItem IsNot Nothing AndAlso StatusItem.HoverBackColor <> Color.Empty Then HoverColor = StatusItem.HoverBackColor

        If HoverColor = Color.Empty Then
            MyBase.OnRenderMenuItemBackground(e)
            Return
        End If

        Dim Bounds As New Rectangle(Point.Empty, e.Item.Size)
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return

        Using Brush As New SolidBrush(HoverColor)
            e.Graphics.FillRectangle(Brush, Bounds)
        End Using

        If _Owner.HoverBorderColor <> Color.Empty AndAlso Bounds.Width > 1 AndAlso Bounds.Height > 1 Then
            Using Pen As New Pen(_Owner.HoverBorderColor)
                e.Graphics.DrawRectangle(Pen, Bounds.X, Bounds.Y, Bounds.Width - 1, Bounds.Height - 1)
            End Using
        End If
    End Sub

    ''' <summary>
    ''' Renders the selected-item indicator inside the native check/image area.
    ''' </summary>
    Protected Overrides Sub OnRenderItemCheck(e As ToolStripItemImageRenderEventArgs)
        If Not _Owner.ShowSelectedCheckMark Then Return

        Dim Bounds = e.ImageRectangle
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return

        Dim OldSmoothingMode = e.Graphics.SmoothingMode
        Try
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Using Pen As New Pen(_Owner.SelectionIndicatorColor, _Owner.SelectionIndicatorThickness)
                Pen.StartCap = LineCap.Round
                Pen.EndCap = LineCap.Round
                Pen.LineJoin = LineJoin.Round

                Dim P1 As New PointF(Bounds.Left + Bounds.Width * 0.2F, Bounds.Top + Bounds.Height * 0.52F)
                Dim P2 As New PointF(Bounds.Left + Bounds.Width * 0.43F, Bounds.Top + Bounds.Height * 0.74F)
                Dim P3 As New PointF(Bounds.Left + Bounds.Width * 0.82F, Bounds.Top + Bounds.Height * 0.28F)

                e.Graphics.DrawLines(Pen, New PointF() {P1, P2, P3})
            End Using
        Finally
            e.Graphics.SmoothingMode = OldSmoothingMode
        End Try
    End Sub

    ''' <summary>
    ''' Supplies StatusSelector colors to the native professional renderer.
    ''' </summary>
    Private NotInheritable Class StatusSelectorColorTable
        Inherits ProfessionalColorTable

        Private ReadOnly _Owner As StatusSelector

        Public Sub New(owner As StatusSelector)
            ArgumentNullException.ThrowIfNull(owner)
            _Owner = owner
            UseSystemColors = False
        End Sub

        Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
            Get
                Return _Owner.MenuBackColor
            End Get
        End Property

        Public Overrides ReadOnly Property MenuBorder As Color
            Get
                Return _Owner.MenuBorderColor
            End Get
        End Property

        Public Overrides ReadOnly Property MenuItemSelected As Color
            Get
                Return _Owner.HoverBackColor
            End Get
        End Property

        Public Overrides ReadOnly Property MenuItemSelectedGradientBegin As Color
            Get
                Return _Owner.HoverBackColor
            End Get
        End Property

        Public Overrides ReadOnly Property MenuItemSelectedGradientEnd As Color
            Get
                Return _Owner.HoverBackColor
            End Get
        End Property

        Public Overrides ReadOnly Property MenuItemBorder As Color
            Get
                If _Owner.HoverBorderColor = Color.Empty Then Return Color.Transparent
                Return _Owner.HoverBorderColor
            End Get
        End Property

        Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
            Get
                Return _Owner.ImageMarginBackColor
            End Get
        End Property

        Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
            Get
                Return _Owner.ImageMarginBackColor
            End Get
        End Property

        Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
            Get
                Return _Owner.ImageMarginBackColor
            End Get
        End Property
    End Class
End Class