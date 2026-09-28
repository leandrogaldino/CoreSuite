Imports System.ComponentModel

''' <summary>
''' Class responsible for smoothly resizing controls or forms.
''' </summary>
Public Class FluidResizer
    Private Const AnimationDuration As Integer = 150
    Private _OldSize As Size
    Private _StartSize As Size
    Private _TargetSize As Size
    Private _AnimationStarted As Long
    Private ReadOnly _Control As Control
    Private ReadOnly _ResizeTimer As Timer

    ''' <summary>
    ''' Occurs when a fluid resizing operation reaches its target size and finishes animating.
    ''' </summary>
    <Category("FluidResizer")>
    <Description("Occurs when a fluid resizing operation reaches its target size and finishes animating.")>
    Public Event ResizeEnd As EventHandler(Of ResizeEndEventArgs)

    ''' <summary>
    ''' Initializes a new instance of the <see cref="FluidResizer"/> class for the specified control.
    ''' </summary>
    ''' <param name="Control">The control or form to be resized.</param>
    Public Sub New(Control As Control)
        ArgumentNullException.ThrowIfNull(Control)
        _Control = Control
        _ResizeTimer = New Timer With {.Interval = 15}
        AddHandler _ResizeTimer.Tick, AddressOf ResizeTimer_Tick
    End Sub

    ''' <summary>
    ''' Smoothly resizes the control to the specified target size.
    ''' </summary>
    ''' <param name="TargetSize">The target size the control should reach.</param>
    ''' <remarks>
    ''' The animation uses a fixed duration and automatically redirects from the current
    ''' size when a new target is requested before the previous animation finishes.
    ''' </remarks>
    Public Sub ResizeTo(TargetSize As Size)
        _OldSize = _Control.Size
        _StartSize = _Control.Size
        _TargetSize = TargetSize
        _AnimationStarted = Environment.TickCount64

        If _Control.Size = _TargetSize Then
            _ResizeTimer.Stop()
            OnResizeEnd(New ResizeEndEventArgs(_OldSize, _TargetSize))
            Return
        End If

        If Not _ResizeTimer.Enabled Then _ResizeTimer.Start()
    End Sub

    ''' <summary>
    ''' Handles the resize timer's tick event and interpolates the control size toward the target size.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">An <see cref="EventArgs"/> that contains no event data.</param>
    Private Sub ResizeTimer_Tick(sender As Object, e As EventArgs)
        Dim Progress = Math.Min(1.0, (Environment.TickCount64 - _AnimationStarted) / CDbl(AnimationDuration))
        Dim EasedProgress = 1.0 - Math.Pow(1.0 - Progress, 3)
        Dim Width = CInt(_StartSize.Width + (_TargetSize.Width - _StartSize.Width) * EasedProgress)
        Dim Height = CInt(_StartSize.Height + (_TargetSize.Height - _StartSize.Height) * EasedProgress)

        _Control.Size = New Size(Width, Height)

        If Progress >= 1.0 Then
            _Control.Size = _TargetSize
            _ResizeTimer.Stop()
            OnResizeEnd(New ResizeEndEventArgs(_OldSize, _TargetSize))
        End If
    End Sub

    ''' <summary>
    ''' Raises the <see cref="ResizeEnd"/> event when resizing completes.
    ''' </summary>
    ''' <param name="e">A <see cref="ResizeEndEventArgs"/> containing original and target sizes.</param>
    Protected Overridable Sub OnResizeEnd(e As ResizeEndEventArgs)
        RaiseEvent ResizeEnd(Me, e)
    End Sub
End Class