Imports System.Drawing
Imports System.Windows.Forms

Public Partial Class ControlSlider
    Private _SnapTimer As Timer
    Private ReadOnly _SnapWatch As New Stopwatch()
    Private _SnapStart As Point
    Private _SnapTarget As Point
    Private _SnapDragStart As Point

    Private ReadOnly Property IsSnapAnimating As Boolean
        Get
            Return _SnapTimer IsNot Nothing AndAlso _SnapTimer.Enabled
        End Get
    End Property

    Private Sub StartSnapAnimation(Target As Point, OriginalStart As Point)
        CancelSnapAnimation()
        _SnapStart = _Child.Location
        _SnapTarget = Target
        _SnapDragStart = OriginalStart
        If _SnapTimer Is Nothing Then
            _SnapTimer = New Timer() With {.Interval = 15}
            AddHandler _SnapTimer.Tick, AddressOf SnapTimerTick
        End If
        _SnapWatch.Restart()
        _SnapTimer.Start()
    End Sub

    Private Sub SnapTimerTick(sender As Object, e As EventArgs)
        If _Child Is Nothing OrElse _Child.IsDisposed Then
            CancelSnapAnimation()
            Return
        End If

        Dim Progress As Double = Math.Min(1.0R, _SnapWatch.Elapsed.TotalMilliseconds / _AnimationDuration)
        Dim Eased As Double = 1.0R - Math.Pow(1.0R - Progress, 3.0R)
        Dim X As Integer = Interpolate(_SnapStart.X, _SnapTarget.X, Eased)
        Dim Y As Integer = Interpolate(_SnapStart.Y, _SnapTarget.Y, Eased)
        _Child.Location = New Point(X, Y)
        If Progress >= 1.0R Then FinishSnapAnimation()
    End Sub

    Private Shared Function Interpolate(StartValue As Integer, EndValue As Integer, T As Double) As Integer
        Return CInt(Math.Round(StartValue + (CDbl(EndValue) - StartValue) * T))
    End Function

    Private Sub FinishSnapAnimation()
        If Not IsSnapAnimating Then Return
        Dim OriginalStart As Point = _SnapDragStart
        CancelSnapAnimation()
        If _Child IsNot Nothing AndAlso Not _Child.IsDisposed Then
            _Child.Location = _SnapTarget
            OnDragCompleted(OriginalStart)
        End If
    End Sub

    Private Sub CancelSnapAnimation()
        If _SnapTimer IsNot Nothing Then _SnapTimer.Stop()
        _SnapWatch.Reset()
    End Sub

    Private Sub DisposeAnimationResources()
        CancelSnapAnimation()
        If _SnapTimer IsNot Nothing Then
            RemoveHandler _SnapTimer.Tick, AddressOf SnapTimerTick
            _SnapTimer.Dispose()
            _SnapTimer = Nothing
        End If
    End Sub
End Class
