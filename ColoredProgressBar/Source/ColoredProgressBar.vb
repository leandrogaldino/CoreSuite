Imports System.ComponentModel
Imports System.Drawing.Drawing2D
''' <summary>
''' Represents a customizable progress bar control that displays progress
''' using a gradient fill based on the current <see cref="Value"/> within the
''' range defined by <see cref="Minimum"/> and <see cref="Maximum"/>.
''' </summary>
Public Class ColoredProgressBar
    Inherits UserControl
    Private _Minimum As Integer = 0
    Private _Maximum As Integer = 100
    Private _Value As Integer = 0
    Private _ProgressTopColor As Color = Color.ForestGreen
    Private _ProgressBottomColor As Color = Color.ForestGreen
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ColoredProgressBar"/> class.
    ''' </summary>
    Public Sub New()
        Size = New Size(150, 24)
        BackColor = Color.WhiteSmoke
    End Sub
    ''' <summary>
    ''' Gets or sets the lowest possible value of the progress range.
    ''' </summary>
    ''' <value>
    ''' The minimum value accepted by the progress bar. The default value is 0.
    ''' </value>
    <Category("ColoredProgressBar")>
    <Description("Defines the lowest possible value of the progress range.")>
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
    ''' Gets or sets the highest possible value of the progress range.
    ''' </summary>
    ''' <value>
    ''' The maximum value accepted by the progress bar. The default value is 100.
    ''' </value>
    <Category("ColoredProgressBar")>
    <Description("Defines the highest possible value of the progress range.")>
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
    ''' Gets or sets the starting color used to draw the progress gradient.
    ''' </summary>
    ''' <value>
    ''' The color applied to the beginning of the gradient fill.
    ''' </value>
    <Category("ColoredProgressBar")>
    <Description("Defines the starting color of the progress gradient fill.")>
    Public Property ProgressStartColor As Color
        Get
            Return _ProgressTopColor
        End Get
        Set(value As Color)
            _ProgressTopColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the ending color used to draw the progress gradient.
    ''' </summary>
    ''' <value>
    ''' The color applied to the end of the gradient fill.
    ''' </value>
    <Category("ColoredProgressBar")>
    <Description("Defines the ending color of the progress gradient fill.")>
    Public Property ProgressEndColor As Color
        Get
            Return _ProgressBottomColor
        End Get
        Set(value As Color)
            _ProgressBottomColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the current progress value displayed by the control.
    ''' </summary>
    ''' <value>
    ''' The current value of the progress bar. The value is automatically
    ''' constrained between <see cref="Minimum"/> and <see cref="Maximum"/>.
    ''' </value>
    <Category("ColoredProgressBar")>
    <Description("Defines the current progress value within the specified range.")>
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
    ''' Paints the control surface and renders the progress gradient according
    ''' to the current <see cref="Value"/>.
    ''' </summary>
    ''' <param name="e">
    ''' Provides data required for painting the control.
    ''' </param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g As Graphics = e.Graphics
        Dim Rect As Rectangle = ClientRectangle
        Dim Brush As New LinearGradientBrush(Rect, ProgressStartColor, ProgressEndColor, LinearGradientMode.Vertical)
        Dim Percent As Single = (_Value - _Minimum) / (_Maximum - _Minimum)
        Rect.Width = CInt(Rect.Width * Percent)
        g.FillRectangle(Brush, Rect)
        Brush.Dispose()
    End Sub
    ''' <summary>
    ''' Refreshes the control rendering when its size changes.
    ''' </summary>
    ''' <param name="e">
    ''' Provides event data for the resize operation.
    ''' </param>
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Invalidate()
    End Sub
End Class