Imports System.Drawing.Drawing2D
Friend NotInheritable Class NavigationButton
    Inherits Button
    Private _PageImage As Image
    Private _ImageSize As Size = New Size(20, 20)
    Private _IsSelected As Boolean
    Private _MouseOver As Boolean
    Private _MousePressed As Boolean
    Private _NormalBackColor As Color = SystemColors.Control
    Private _HoverBackColor As Color = SystemColors.ControlLight
    Private _SelectedBackColor As Color = SystemColors.Highlight
    Private _NormalForeColor As Color = SystemColors.ControlText
    Private _SelectedForeColor As Color = SystemColors.HighlightText
    Private _IndicatorColor As Color = SystemColors.Highlight
    Private _IndicatorWidth As Integer = 4
    Private _IndicatorOnRight As Boolean
    Public Sub New()
        FlatStyle = FlatStyle.Flat
        FlatAppearance.BorderSize = 0
        UseVisualStyleBackColor = False
        TextAlign = ContentAlignment.MiddleLeft
        ImageAlign = ContentAlignment.MiddleLeft
        TextImageRelation = TextImageRelation.ImageBeforeText
        AutoEllipsis = True
        Cursor = Cursors.Hand
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
    End Sub
    Public Property Page As NavigationPage
    Public Property PageImage As Image
        Get
            Return _PageImage
        End Get
        Set(Value As Image)
            If ReferenceEquals(_PageImage, Value) Then Return
            _PageImage = Value
            Invalidate()
        End Set
    End Property
    Public Property ImageSize As Size
        Get
            Return _ImageSize
        End Get
        Set(Value As Size)
            _ImageSize = Value
            Invalidate()
        End Set
    End Property
    Public Property IsSelected As Boolean
        Get
            Return _IsSelected
        End Get
        Set(Value As Boolean)
            If _IsSelected = Value Then Return
            _IsSelected = Value
            Invalidate()
        End Set
    End Property
    Public Property NormalBackColor As Color
        Get
            Return _NormalBackColor
        End Get
        Set(Value As Color)
            _NormalBackColor = Value
            Invalidate()
        End Set
    End Property
    Public Property HoverBackColor As Color
        Get
            Return _HoverBackColor
        End Get
        Set(Value As Color)
            _HoverBackColor = Value
            Invalidate()
        End Set
    End Property
    Public Property SelectedBackColor As Color
        Get
            Return _SelectedBackColor
        End Get
        Set(Value As Color)
            _SelectedBackColor = Value
            Invalidate()
        End Set
    End Property
    Public Property NormalForeColor As Color
        Get
            Return _NormalForeColor
        End Get
        Set(Value As Color)
            _NormalForeColor = Value
            Invalidate()
        End Set
    End Property
    Public Property SelectedForeColor As Color
        Get
            Return _SelectedForeColor
        End Get
        Set(Value As Color)
            _SelectedForeColor = Value
            Invalidate()
        End Set
    End Property
    Public Property IndicatorColor As Color
        Get
            Return _IndicatorColor
        End Get
        Set(Value As Color)
            _IndicatorColor = Value
            Invalidate()
        End Set
    End Property
    Public Property IndicatorWidth As Integer
        Get
            Return _IndicatorWidth
        End Get
        Set(Value As Integer)
            _IndicatorWidth = Value
            Invalidate()
        End Set
    End Property
    Public Property IndicatorOnRight As Boolean
        Get
            Return _IndicatorOnRight
        End Get
        Set(Value As Boolean)
            _IndicatorOnRight = Value
            Invalidate()
        End Set
    End Property
    Protected Overrides Sub OnMouseEnter(E As EventArgs)
        MyBase.OnMouseEnter(E)
        _MouseOver = True
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseLeave(E As EventArgs)
        MyBase.OnMouseLeave(E)
        _MouseOver = False
        _MousePressed = False
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseDown(Mevent As MouseEventArgs)
        MyBase.OnMouseDown(Mevent)
        If Mevent.Button = MouseButtons.Left Then
            _MousePressed = True
            Invalidate()
        End If
    End Sub
    Protected Overrides Sub OnMouseUp(Mevent As MouseEventArgs)
        MyBase.OnMouseUp(Mevent)
        _MousePressed = False
        Invalidate()
    End Sub
    Protected Overrides Sub OnEnabledChanged(E As EventArgs)
        MyBase.OnEnabledChanged(E)
        Invalidate()
    End Sub
    Protected Overrides Sub OnPaint(E As PaintEventArgs)
        Dim Background As Color = ResolveBackgroundColor()
        Dim Foreground As Color = If(_IsSelected, _SelectedForeColor, _NormalForeColor)
        If Not Enabled Then Foreground = SystemColors.GrayText
        Using BackgroundBrush As New SolidBrush(Background)
            E.Graphics.FillRectangle(BackgroundBrush, ClientRectangle)
        End Using
        If _IsSelected AndAlso _IndicatorWidth > 0 Then
            Dim IndicatorBounds As Rectangle
            If _IndicatorOnRight Then
                IndicatorBounds = New Rectangle(Math.Max(0, Width - _IndicatorWidth), 0, _IndicatorWidth, Height)
            Else
                IndicatorBounds = New Rectangle(0, 0, _IndicatorWidth, Height)
            End If
            Using IndicatorBrush As New SolidBrush(_IndicatorColor)
                E.Graphics.FillRectangle(IndicatorBrush, IndicatorBounds)
            End Using
        End If
        Dim ContentBounds As New Rectangle(Padding.Left, Padding.Top, Math.Max(0, Width - Padding.Horizontal), Math.Max(0, Height - Padding.Vertical))
        Dim DrawImage As Boolean = _PageImage IsNot Nothing AndAlso _ImageSize.Width > 0 AndAlso _ImageSize.Height > 0
        Dim TextBounds As Rectangle = ContentBounds
        If DrawImage Then
            Dim ImageWidth As Integer = Math.Min(_ImageSize.Width, ContentBounds.Width)
            Dim ImageHeight As Integer = Math.Min(_ImageSize.Height, ContentBounds.Height)
            Dim ImageY As Integer = ContentBounds.Top + Math.Max(0, (ContentBounds.Height - ImageHeight) \ 2)
            Dim ImageBounds As Rectangle
            If Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes Then
                ImageBounds = New Rectangle(Math.Max(ContentBounds.Left, ContentBounds.Right - ImageWidth), ImageY, ImageWidth, ImageHeight)
                TextBounds.Width = Math.Max(0, ImageBounds.Left - ContentBounds.Left - 8)
            Else
                ImageBounds = New Rectangle(ContentBounds.Left, ImageY, ImageWidth, ImageHeight)
                TextBounds.X = Math.Min(ContentBounds.Right, ImageBounds.Right + 8)
                TextBounds.Width = Math.Max(0, ContentBounds.Right - TextBounds.X)
            End If
            E.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
            E.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
            E.Graphics.DrawImage(_PageImage, ImageBounds)
        End If
        Dim Flags As TextFormatFlags = TextFormatFlags.VerticalCenter Or TextFormatFlags.SingleLine Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix
        If Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes Then Flags = Flags Or TextFormatFlags.RightToLeft Or TextFormatFlags.Right
        TextRenderer.DrawText(E.Graphics, Text, Font, TextBounds, Foreground, Flags)
        If Focused AndAlso ShowFocusCues Then
            Dim FocusBounds As Rectangle = Rectangle.Inflate(ClientRectangle, -3, -3)
            ControlPaint.DrawFocusRectangle(E.Graphics, FocusBounds, Foreground, Background)
        End If
    End Sub
    Private Function ResolveBackgroundColor() As Color
        If _IsSelected Then Return _SelectedBackColor
        If Enabled AndAlso (_MouseOver OrElse _MousePressed) Then Return _HoverBackColor
        Return _NormalBackColor
    End Function
End Class
