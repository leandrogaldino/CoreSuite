Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' A custom progress bar that fills with a gradient color based on the <see cref="Value"/>.
''' </summary>
Public Class ColoredProgressBar
    Inherits UserControl

    Private _Minimum As Integer = 0
    Private _Maximum As Integer = 100
    Private _Value As Integer = 0
    Private _ProgressTopColor As Color = Color.ForestGreen
    Private _ProgressBottomColor As Color = Color.ForestGreen

    Public Sub New()
        Size = New Size(150, 24)
        BackColor = Color.WhiteSmoke
    End Sub

    ''' <summary>
    ''' Gets or sets the minimum value of the progress bar.
    ''' </summary>
    <Browsable(True)>
    Public Property Minimum As Integer
        Get
            Return _Minimum
        End Get
        Set(value As Integer)
            If value < 0 Then value = 0
            If value > _Maximum Then _Maximum = value
            _Minimum = value
            If _Value < _Minimum Then _Value = _Minimum
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the maximum value of the progress bar.
    ''' </summary>
    <Browsable(True)>
    Public Property Maximum As Integer
        Get
            Return _Maximum
        End Get
        Set(value As Integer)
            If value < _Minimum Then _Minimum = value
            _Maximum = value
            If _Value > _Maximum Then _Value = _Maximum
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the top color of the gradient fill.
    ''' </summary>
    <Browsable(True)>
    Public Property ProgressTopColor As Color
        Get
            Return _ProgressTopColor
        End Get
        Set(value As Color)
            _ProgressTopColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the bottom color of the gradient fill.
    ''' </summary>
    <Browsable(True)>
    Public Property ProgressBottomColor As Color
        Get
            Return _ProgressBottomColor
        End Get
        Set(value As Color)
            _ProgressBottomColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the current value of the progress bar.
    ''' </summary>
    <Browsable(True)>
    Public Property Value As Integer
        Get
            Return _Value
        End Get
        Set(value As Integer)
            Dim oldValue As Integer = _Value
            If value < _Minimum Then
                _Value = _Minimum
            ElseIf value > _Maximum Then
                _Value = _Maximum
            Else
                _Value = value
            End If
            Dim NewValueRect As Rectangle = ClientRectangle
            Dim OldValueRect As Rectangle = ClientRectangle
            Dim Percent As Single
            Percent = (_Value - _Minimum) / (_Maximum - _Minimum)
            NewValueRect.Width = CInt(ClientRectangle.Width * Percent)
            Percent = (oldValue - _Minimum) / (_Maximum - _Minimum)
            OldValueRect.Width = CInt(ClientRectangle.Width * Percent)
            Dim UpdateRect As New Rectangle()
            If NewValueRect.Width > OldValueRect.Width Then
                UpdateRect.X = OldValueRect.Width
                UpdateRect.Width = NewValueRect.Width - OldValueRect.Width
            Else
                UpdateRect.X = NewValueRect.Width
                UpdateRect.Width = OldValueRect.Width - NewValueRect.Width
            End If
            UpdateRect.Height = Me.Height
            Invalidate(UpdateRect)
        End Set
    End Property

    ''' <summary>
    ''' Paints the progress bar with the gradient fill according to the current <see cref="Value"/>.
    ''' </summary>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g As Graphics = e.Graphics
        Dim Rect As Rectangle = ClientRectangle
        Dim Brush As New LinearGradientBrush(Rect, ProgressTopColor, ProgressBottomColor, LinearGradientMode.Vertical)
        Dim Percent As Single = (_Value - _Minimum) / (_Maximum - _Minimum)
        Rect.Width = CInt(Rect.Width * Percent)
        g.FillRectangle(Brush, Rect)
        Brush.Dispose()
    End Sub

    ''' <summary>
    ''' Refreshes the control when resized.
    ''' </summary>
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Invalidate()
    End Sub

End Class