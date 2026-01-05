Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Public Class SplitButton
    Inherits NoFocusCueButton
    Private _BmpSplit As Bitmap
    Private _IsDropDown As Boolean

    '<DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property ButtonMenuStrip As ContextMenuStrip

    ''' <summary>
    ''' Especifica se o controle será tratado como um DropDownButton.
    ''' </summary>
    ''' <returns>Verdadeiro para DropDown e Falso para SplitButton</returns>
    <Category("Comportamento")>
    <DefaultValue(False)>
    Public Property IsDropDown As Boolean
        Get
            Return _IsDropDown
        End Get
        Set(value As Boolean)
            _IsDropDown = value
            Split(Color.Gainsboro)
        End Set
    End Property
    Public Sub New()
        Width = 125
        Split(Color.Gainsboro)

    End Sub
    Protected Overrides Sub InitLayout()
        MyBase.InitLayout()
        Split(Color.Gainsboro)
    End Sub
    Protected Overrides Sub OnClick(e As EventArgs)
        If Position(IsDropDown) Then
            If ButtonMenuStrip IsNot Nothing Then ButtonMenuStrip.Show(Me, 0, Height)
        Else
            MyBase.OnClick(e)
        End If
    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        Split(Color.Gainsboro)
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        Split(Color.Silver)
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        Split(Color.Gainsboro)
    End Sub
    Private Function Position(isdropdown As Boolean) As Boolean
        Dim Pos As Boolean = False
        Dim X As Integer = PointToClient(MousePosition).X
        Dim y As Integer = PointToClient(MousePosition).Y
        If Not Me.IsDropDown Then
            If X > Width - _BmpSplit.Width - 4 Then
                If X < Width Then
                    If y > 0 Then
                        If y < Height Then
                            Pos = True
                        End If
                    End If
                End If
            End If
        Else
            Pos = True
        End If
        Return Pos
    End Function

    Private Sub Split(LineColor As Color)
        Dim Graphic As Graphics
        Dim ArrowX As Integer
        Dim ArrowY As Integer
        Dim Points As Point()
        _BmpSplit = New Bitmap(25, Height)
        Graphic = Graphics.FromImage(_BmpSplit)
        Graphic.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        ArrowX = _BmpSplit.Width - 16
        ArrowY = _BmpSplit.Height / 2 - 4.5
        Points = {New Point(ArrowX, ArrowY), New Point(ArrowX + 7, ArrowY), New Point(ArrowX + 3, ArrowY + 4)}
        Graphic.FillPolygon(New SolidBrush(Color.DimGray), Points)
        If Not IsDropDown Then Graphic.DrawLine(New Pen(LineColor), New Point(0, (_BmpSplit.Height - 7) / 4), New PointF(0, (_BmpSplit.Height - 7) - (_BmpSplit.Height - 7) / 4))
        ImageAlign = ContentAlignment.MiddleRight
        TextImageRelation = TextImageRelation.TextBeforeImage
        Image = _BmpSplit
    End Sub

End Class
