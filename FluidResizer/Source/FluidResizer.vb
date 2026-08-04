Imports System.ComponentModel

''' <summary>
''' Class responsible for smoothly resizing controls or forms.
''' </summary>
Public Class FluidResizer
    Private _OldSize As Size
    Private _TargetSize As Size
    Private _IsResizing As Boolean = False
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
        _Control = Control
        _ResizeTimer = New Timer With {.Interval = 1}
        AddHandler _ResizeTimer.Tick, AddressOf ResizeTimer_Tick
    End Sub
    ''' <summary>
    ''' Sets the desired final size for the control and starts the smooth resizing process.
    ''' </summary>
    ''' <param name="TargetSize">The target size the control should reach.</param>
    ''' <remarks>
    ''' The resizing is performed gradually in steps, with increments calculated
    ''' to ensure a smooth transition.
    ''' </remarks>
    Public Sub SetSize(TargetSize As Size)
        If _IsResizing Then Return
        _OldSize = _Control.Size
        _TargetSize = TargetSize
        _IsResizing = True
        If Not _ResizeTimer.Enabled Then
            _ResizeTimer.Start()
        End If
    End Sub
    ''' <summary>
    ''' Handles the resize timer's tick event, incrementally moving the control's
    ''' width and height toward the target size until it is reached.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">An <see cref="EventArgs"/> that contains no event data.</param>
    Private Sub ResizeTimer_Tick(sender As Object, e As EventArgs)
        Dim StepWidth As Integer = Math.Max(1, Math.Abs(_TargetSize.Width - _Control.Width) / 5)
        Dim StepHeight As Integer = Math.Max(1, Math.Abs(_TargetSize.Height - _Control.Height) / 5)
        If _Control.Width < _TargetSize.Width Then
            _Control.Width = Math.Min(_Control.Width + StepWidth, _TargetSize.Width)
        ElseIf _Control.Width > _TargetSize.Width Then
            _Control.Width = Math.Max(_Control.Width - StepWidth, _TargetSize.Width)
        End If
        If _Control.Height < _TargetSize.Height Then
            _Control.Height = Math.Min(_Control.Height + StepHeight, _TargetSize.Height)
        ElseIf _Control.Height > _TargetSize.Height Then
            _Control.Height = Math.Max(_Control.Height - StepHeight, _TargetSize.Height)
        End If
        If _Control.Width = _TargetSize.Width AndAlso _Control.Height = _TargetSize.Height Then
            _ResizeTimer.Stop()
            _IsResizing = False
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