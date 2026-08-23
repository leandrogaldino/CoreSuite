Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Friend NotInheritable Class StatusSelectorColorTable
    Inherits ProfessionalColorTable
    Private ReadOnly _Owner As StatusSelector
    Public Sub New(owner As StatusSelector)
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
    Public Overrides ReadOnly Property MenuItemBorder As Color
        Get
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

Friend NotInheritable Class StatusSelectorRenderer
    Inherits ToolStripProfessionalRenderer
    Private ReadOnly _Owner As StatusSelector
    Public Sub New(owner As StatusSelector)
        MyBase.New(New StatusSelectorColorTable(owner))
        _Owner = owner
        RoundedEdges = owner.MenuRoundedEdges
    End Sub
    Protected Overrides Sub OnRenderMenuItemBackground(e As ToolStripItemRenderEventArgs)
        Dim MenuItem = TryCast(e.Item, ToolStripMenuItem)
        Dim StatusItem = TryCast(MenuItem?.Tag, StatusSelectorItem)
        If StatusItem Is Nothing Then
            MyBase.OnRenderMenuItemBackground(e)
            Return
        End If
        Dim BackColor = ResolveBackColor(StatusItem, e.Item.Selected)
        Using Brush As New SolidBrush(BackColor)
            e.Graphics.FillRectangle(Brush, New Rectangle(Point.Empty, e.Item.Size))
        End Using
        If e.Item.Selected AndAlso _Owner.HoverBorderColor <> Color.Empty Then
            Using Pen As New Pen(_Owner.HoverBorderColor)
                Dim Bounds = New Rectangle(0, 0, Math.Max(0, e.Item.Width - 1), Math.Max(0, e.Item.Height - 1))
                e.Graphics.DrawRectangle(Pen, Bounds)
            End Using
        End If
    End Sub
    Protected Overrides Sub OnRenderItemText(e As ToolStripItemTextRenderEventArgs)
        Dim MenuItem = TryCast(e.Item, ToolStripMenuItem)
        Dim StatusItem = TryCast(MenuItem?.Tag, StatusSelectorItem)
        If StatusItem IsNot Nothing Then
            e.TextColor = ResolveForeColor(StatusItem, e.Item.Selected)
        End If
        MyBase.OnRenderItemText(e)
    End Sub
    Protected Overrides Sub OnRenderItemCheck(e As ToolStripItemImageRenderEventArgs)
        Dim MenuItem = TryCast(e.Item, ToolStripMenuItem)
        If MenuItem Is Nothing OrElse Not MenuItem.Checked OrElse Not _Owner.ShowSelectedCheckMark Then Return
        Dim Rect = e.ImageRectangle
        If Rect.Width <= 0 OrElse Rect.Height <= 0 Then Return
        Dim X1 = Rect.Left + CInt(Rect.Width * 0.2)
        Dim Y1 = Rect.Top + CInt(Rect.Height * 0.55)
        Dim X2 = Rect.Left + CInt(Rect.Width * 0.42)
        Dim Y2 = Rect.Top + CInt(Rect.Height * 0.75)
        Dim X3 = Rect.Left + CInt(Rect.Width * 0.82)
        Dim Y3 = Rect.Top + CInt(Rect.Height * 0.28)
        Using Pen As New Pen(_Owner.SelectionIndicatorColor, Math.Max(1.5F, _Owner.SelectionIndicatorThickness)) With {.StartCap = LineCap.Round, .EndCap = LineCap.Round, .LineJoin = LineJoin.Round}
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            e.Graphics.DrawLines(Pen, {New Point(X1, Y1), New Point(X2, Y2), New Point(X3, Y3)})
        End Using
    End Sub
    Private Function ResolveBackColor(item As StatusSelectorItem, selected As Boolean) As Color
        If selected Then
            If item.HoverBackColor <> Color.Empty Then Return item.HoverBackColor
            Return _Owner.HoverBackColor
        End If
        If item.BackColor <> Color.Empty Then Return item.BackColor
        Return _Owner.MenuBackColor
    End Function
    Private Function ResolveForeColor(item As StatusSelectorItem, selected As Boolean) As Color
        If selected Then
            If item.HoverForeColor <> Color.Empty Then Return item.HoverForeColor
            If _Owner.HoverForeColor <> Color.Empty Then Return _Owner.HoverForeColor
        End If
        Return item.ForeColor
    End Function
End Class
