Imports System.ComponentModel
''' <summary>
''' Represents a customizable horizontal or vertical separator line.
''' </summary>
<DefaultProperty("SeparatorColor")>
Public Class Separator
    Inherits Control
    Private _SeparatorColor As Color = SystemColors.ControlDark
    Private _Thickness As Integer = 1
    Private _Orientation As Orientation = Orientation.Horizontal
    Private _SeparatorAlignment As SeparatorAlignment = SeparatorAlignment.Center
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property Text As String
        Get
            Return String.Empty
        End Get
        Set(value As String)
        End Set
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Shadows Event TextChanged As EventHandler
    ''' <summary>
    ''' Gets or sets the color used to draw the separator.
    ''' </summary>
    <Category("Separator")>
    <DefaultValue(GetType(Color), "ControlDark")>
    <Description("Gets or sets the color used to draw the separator.")>
    Public Property SeparatorColor As Color
        Get
            Return _SeparatorColor
        End Get
        Set(value As Color)
            If _SeparatorColor = value Then Return
            _SeparatorColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the thickness of the separator.
    ''' </summary>
    <Category("Separator")>
    <DefaultValue(1)>
    <Description("Gets or sets the thickness of the separator.")>
    Public Property Thickness As Integer
        Get
            Return _Thickness
        End Get
        Set(value As Integer)
            value = Math.Max(1, value)
            If _Thickness = value Then Return
            _Thickness = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the separator is displayed horizontally or vertically.
    ''' </summary>
    <Category("Separator")>
    <DefaultValue(GetType(Orientation), "Horizontal")>
    <Description("Gets or sets whether the separator is displayed horizontally or vertically.")>
    Public Property Orientation As Orientation
        Get
            Return _Orientation
        End Get
        Set(value As Orientation)
            If _Orientation = value Then Return
            _Orientation = value
            If value = Orientation.Horizontal AndAlso Width < Height Then
                Size = New Size(Height, Width)
            ElseIf value = Orientation.Vertical AndAlso Height < Width Then
                Size = New Size(Height, Width)
            End If
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the alignment of the separator line within the control bounds.
    ''' </summary>
    <Category("Separator")>
    <DefaultValue(GetType(SeparatorAlignment), "Center")>
    <Description("Gets or sets the alignment of the separator line within the control bounds.")>
    Public Property SeparatorAlignment As SeparatorAlignment
        Get
            Return _SeparatorAlignment
        End Get
        Set(value As SeparatorAlignment)
            If _SeparatorAlignment = value Then Return
            _SeparatorAlignment = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Separator"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor, True)
        SetStyle(ControlStyles.Selectable, False)
        TabStop = False
        BackColor = Color.Transparent
        Margin = New Padding(0)
        Padding = New Padding(0)
        AccessibleRole = AccessibleRole.Separator
    End Sub
    ''' <summary>
    ''' Paints the separator according to its configured orientation, alignment, color, and thickness.
    ''' </summary>
    ''' <param name="e">The event data containing the graphics surface used for painting.</param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Using Brush As New SolidBrush(_SeparatorColor)
            If _Orientation = Orientation.Horizontal Then
                Dim AvailableWidth = Math.Max(0, ClientSize.Width - Padding.Horizontal)
                Dim AvailableHeight = Math.Max(0, ClientSize.Height - Padding.Vertical)
                Dim LineThickness = Math.Min(_Thickness, AvailableHeight)
                If AvailableWidth = 0 OrElse LineThickness = 0 Then Return
                Dim Y = GetAlignedPosition(Padding.Top, AvailableHeight, LineThickness)
                e.Graphics.FillRectangle(Brush, Padding.Left, Y, AvailableWidth, LineThickness)
            Else
                Dim AvailableWidth = Math.Max(0, ClientSize.Width - Padding.Horizontal)
                Dim AvailableHeight = Math.Max(0, ClientSize.Height - Padding.Vertical)
                Dim LineThickness = Math.Min(_Thickness, AvailableWidth)
                If AvailableHeight = 0 OrElse LineThickness = 0 Then Return
                Dim X = GetAlignedPosition(Padding.Left, AvailableWidth, LineThickness)
                e.Graphics.FillRectangle(Brush, X, Padding.Top, LineThickness, AvailableHeight)
            End If
        End Using
    End Sub
    ''' <summary>
    ''' Calculates the aligned position of the separator within the available space.
    ''' </summary>
    ''' <param name="StartPosition">The initial position of the available area.</param>
    ''' <param name="AvailableSize">The size of the available area.</param>
    ''' <param name="LineSize">The thickness of the separator line.</param>
    ''' <returns>The calculated position of the separator line.</returns>
    Private Function GetAlignedPosition(StartPosition As Integer, AvailableSize As Integer, LineSize As Integer) As Integer
        Dim RemainingSize = Math.Max(0, AvailableSize - LineSize)
        Select Case _SeparatorAlignment
            Case SeparatorAlignment.Near
                Return StartPosition
            Case SeparatorAlignment.Far
                Return StartPosition + RemainingSize
            Case Else
                Return StartPosition + RemainingSize \ 2
        End Select
    End Function
End Class