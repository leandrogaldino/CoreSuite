
Imports CoreSuite.Helpers
Imports System
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Provides smooth, animated resizing for Windows Forms controls and forms.
''' </summary>
''' <remarks>
''' Supports interruption and redirection of an active animation.
''' All operations must be performed on the associated control's UI thread.
''' </remarks>
Public Class FluidResizer
    Implements IDisposable

    Private Const DefaultAnimationDuration As Integer = 150
    Private Const TimerInterval As Integer = 15

    Private ReadOnly _Control As Control
    Private ReadOnly _ResizeTimer As Timer

    Private _StartSize As Size
    Private _TargetSize As Size
    Private _AnimationStarted As Long
    Private _AnimationVersion As Long
    Private _AnimationDuration As Integer = DefaultAnimationDuration
    Private _Disposed As Boolean

    ''' <summary>
    ''' Occurs when a resizing animation completes successfully.
    ''' </summary>
    Public Event ResizeEnd As EventHandler(Of ResizeEndEventArgs)

    ''' <summary>
    ''' Gets or sets the duration of the resizing animation, in milliseconds.
    ''' </summary>
    ''' <remarks>
    ''' The default duration is 150 milliseconds.
    ''' Changes to this property also affect an animation currently in progress.
    ''' </remarks>
    Public Property AnimationDuration As Integer
        Get
            Return _AnimationDuration
        End Get
        Set(value As Integer)
            ThrowIfDisposed()

            If value <= 0 Then
                Throw New ArgumentOutOfRangeException(
                    NameOf(value),
                    "Animation duration must be greater than zero.")
            End If

            _AnimationDuration = value
        End Set
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="FluidResizer"/> class.
    ''' </summary>
    ''' <param name="Control">
    ''' The control or form whose size will be animated.
    ''' </param>
    Public Sub New(Control As Control)
        If Control Is Nothing Then
            Throw New ArgumentNullException(NameOf(Control))
        End If

        If Control.IsDisposed OrElse Control.Disposing Then
            Throw New ObjectDisposedException(Control.GetType().Name)
        End If

        _Control = Control

        _ResizeTimer = New Timer With {
            .Interval = TimerInterval
        }

        AddHandler _ResizeTimer.Tick, AddressOf ResizeTimer_Tick
        AddHandler _Control.Disposed, AddressOf Control_Disposed
    End Sub

    ''' <summary>
    ''' Smoothly resizes the associated control to the specified target size.
    ''' </summary>
    ''' <param name="TargetSize">
    ''' The desired final size of the control.
    ''' </param>
    ''' <remarks>
    ''' If another animation is in progress, it is interrupted and a new
    ''' animation starts from the control's current size.
    ''' No completion event is raised when the control is already at
    ''' the requested size.
    ''' </remarks>
    Public Sub ResizeTo(TargetSize As Size)
        ThrowIfDisposed()

        If _Control.InvokeRequired Then
            Throw New InvalidOperationException(
                "FluidResizer must be accessed from the control's UI thread.")
        End If

        If TargetSize.Width < 0 OrElse TargetSize.Height < 0 Then
            Throw New ArgumentOutOfRangeException(
                NameOf(TargetSize),
                "Target size dimensions cannot be negative.")
        End If

        _AnimationVersion += 1

        _StartSize = _Control.Size
        _TargetSize = TargetSize
        _AnimationStarted = SystemClockHelper.TickCount64

        If _StartSize = _TargetSize Then
            _ResizeTimer.Stop()
            Return
        End If

        If Not _ResizeTimer.Enabled Then
            _ResizeTimer.Start()
        End If
    End Sub

    ''' <summary>
    ''' Updates the control size using cubic ease-out interpolation.
    ''' </summary>
    Private Sub ResizeTimer_Tick(sender As Object, e As EventArgs)
        If _Disposed Then Return

        If _Control.IsDisposed OrElse _Control.Disposing Then
            Dispose()
            Return
        End If

        Dim CurrentVersion As Long = _AnimationVersion

        Dim Elapsed As Long =
            SystemClockHelper.TickCount64 - _AnimationStarted

        Dim Progress As Double =
            Math.Max(0.0, Math.Min(1.0,
                Elapsed / CDbl(_AnimationDuration)))

        Dim EasedProgress As Double =
            1.0 - Math.Pow(1.0 - Progress, 3)

        Dim NewWidth As Integer = CInt(
            _StartSize.Width +
            (_TargetSize.Width - _StartSize.Width) * EasedProgress)

        Dim NewHeight As Integer = CInt(
            _StartSize.Height +
            (_TargetSize.Height - _StartSize.Height) * EasedProgress)

        Dim NewSize As New Size(NewWidth, NewHeight)

        If _Control.Size <> NewSize Then
            _Control.Size = NewSize
        End If

        ' The resize operation may have triggered another animation
        ' or caused the associated control to be disposed.
        If _Disposed OrElse CurrentVersion <> _AnimationVersion Then
            Return
        End If

        If Progress >= 1.0 Then
            _ResizeTimer.Stop()

            OnResizeEnd(New ResizeEndEventArgs(
                _StartSize,
                _Control.Size))
        End If
    End Sub

    ''' <summary>
    ''' Raises the <see cref="ResizeEnd"/> event.
    ''' </summary>
    ''' <param name="e">
    ''' Event data containing the initial and final control sizes.
    ''' </param>
    Protected Overridable Sub OnResizeEnd(e As ResizeEndEventArgs)
        RaiseEvent ResizeEnd(Me, e)
    End Sub

    ''' <summary>
    ''' Releases resources when the associated control is disposed.
    ''' </summary>
    Private Sub Control_Disposed(sender As Object, e As EventArgs)
        Dispose()
    End Sub

    ''' <summary>
    ''' Throws an exception if the resizer or its associated control
    ''' has already been disposed.
    ''' </summary>
    Private Sub ThrowIfDisposed()
        If _Disposed Then
            Throw New ObjectDisposedException(NameOf(FluidResizer))
        End If

        If _Control.IsDisposed OrElse _Control.Disposing Then
            Throw New ObjectDisposedException(_Control.GetType().Name)
        End If
    End Sub

    ''' <summary>
    ''' Releases all resources used by the resizer.
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ''' <summary>
    ''' Releases managed resources associated with the resizer.
    ''' </summary>
    ''' <param name="Disposing">
    ''' Indicates whether managed resources should be released.
    ''' </param>
    Protected Overridable Sub Dispose(Disposing As Boolean)
        If _Disposed Then Return

        _Disposed = True

        If Disposing Then
            _ResizeTimer.Stop()

            RemoveHandler _ResizeTimer.Tick, AddressOf ResizeTimer_Tick
            RemoveHandler _Control.Disposed, AddressOf Control_Disposed

            _ResizeTimer.Dispose()
        End If
    End Sub

End Class
