Imports System.ComponentModel

Public Partial Class ColoredProgressBar

    ' --------------------------------------------------
    ' Animation properties
    ' --------------------------------------------------

    <Category("Animation")>
    <Description("Enables smooth animated progress transitions.")>
    <DefaultValue(True)>
    Public Property AnimateChanges As Boolean
        Get
            Return _AnimateChanges
        End Get
        Set(value As Boolean)
            If _AnimateChanges = value Then Return

            _AnimateChanges = value

            If Not value Then
                SnapToTarget()
            End If
        End Set
    End Property

    <Category("Animation")>
    <Description("Gets or sets the animation duration in milliseconds. Zero disables transitions.")>
    <DefaultValue(250)>
    Public Property AnimationDuration As Integer
        Get
            Return _AnimationDuration
        End Get
        Set(value As Integer)
            Dim NewDuration As Integer = Math.Max(0, value)

            If _AnimationDuration = NewDuration Then Return

            _AnimationDuration = NewDuration

            If _AnimationTimer.Enabled Then
                UpdateVisualProgress()
            End If
        End Set
    End Property

    ' --------------------------------------------------
    ' Progress animation
    ' --------------------------------------------------

    Private Sub UpdateVisualProgress()
        Dim Target As Double = Percentage / 100.0R

        If Not _AnimateChanges OrElse
           _AnimationDuration = 0 OrElse
           Not IsHandleCreated OrElse
           Not Visible OrElse
           IsDesignTime Then

            SnapToTarget()
            Return
        End If

        If Math.Abs(Target - _DisplayedProgress) < 0.0000001R Then
            SnapToTarget()
            Return
        End If

        ' Continue from the currently displayed position,
        ' even when a previous animation is still running.

        _AnimationStartProgress = _DisplayedProgress
        _AnimationTargetProgress = Target

        _AnimationClock.Restart()
        _AnimationTimer.Start()

        Invalidate()
    End Sub

    Private Sub SnapToTarget()
        StopAnimation()

        _DisplayedProgress = Percentage / 100.0R

        Invalidate()
    End Sub

    Private Sub StopAnimation()
        _AnimationTimer.Stop()
        _AnimationClock.Stop()
    End Sub

    Private ReadOnly Property IsDesignTime As Boolean
        Get
            Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse
                   (Site IsNot Nothing AndAlso Site.DesignMode)
        End Get
    End Property

    ' --------------------------------------------------
    ' Animation engine
    ' --------------------------------------------------

    Private Sub AnimationTimer_Tick(sender As Object, e As EventArgs)
        If _AnimationDuration <= 0 Then
            SnapToTarget()
            Return
        End If

        Dim Elapsed As Double = _AnimationClock.Elapsed.TotalMilliseconds

        Dim Progress As Double =
            Math.Min(1.0R, Elapsed / CDbl(_AnimationDuration))

        ' Cubic Ease-Out:
        ' Starts quickly and gently slows down near the target.

        Dim Remaining As Double = 1.0R - Progress

        Dim EasedProgress As Double =
            1.0R - Remaining * Remaining * Remaining

        _DisplayedProgress =
            _AnimationStartProgress +
            (_AnimationTargetProgress - _AnimationStartProgress) * EasedProgress

        If Progress >= 1.0R Then
            _DisplayedProgress = _AnimationTargetProgress
            StopAnimation()
        End If

        Invalidate()
    End Sub

End Class
