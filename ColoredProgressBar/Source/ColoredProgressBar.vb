
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Diagnostics
Imports System.Windows.Forms

''' <summary>
''' Represents a customizable progress bar with animated transitions,
''' rounded corners, gradient fill, borders and percentage text.
''' </summary>
<DefaultEvent("ValueChanged")>
Public Partial Class ColoredProgressBar
    Inherits Control

    Private _Minimum As Integer = 0
    Private _Maximum As Integer = 100
    Private _Value As Integer = 0
    Private _Step As Integer = 10

    Private _ProgressStartColor As Color = Color.FromArgb(16, 185, 129)
    Private _ProgressEndColor As Color = Color.FromArgb(5, 150, 105)
    Private _TrackColor As Color = Color.FromArgb(229, 231, 235)
    Private _BorderColor As Color = Color.FromArgb(203, 213, 225)
    Private _ProgressTextColor As Color = Color.White

    Private _GradientDirection As LinearGradientMode = LinearGradientMode.Vertical
    Private _CornerRadius As Integer = 9
    Private _BorderThickness As Integer = 0
    Private _ShowPercentage As Boolean = False

    Private _AnimateChanges As Boolean = True
    Private _AnimationDuration As Integer = 250

    Private ReadOnly _AnimationTimer As Timer
    Private ReadOnly _AnimationClock As Stopwatch

    Private _DisplayedProgress As Double = 0.0R
    Private _AnimationStartProgress As Double
    Private _AnimationTargetProgress As Double

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or
                 ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or
                 ControlStyles.SupportsTransparentBackColor, True)

        SetStyle(ControlStyles.Selectable, False)

        BackColor = Color.Transparent
        ForeColor = Color.FromArgb(31, 41, 55)
        TabStop = False

        _AnimationClock = New Stopwatch()
        _AnimationTimer = New Timer() With {
            .Interval = 15
        }

        AddHandler _AnimationTimer.Tick, AddressOf AnimationTimer_Tick

        UpdateStyles()
    End Sub

    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return New Size(180, 24)
        End Get
    End Property

    ' --------------------------------------------------
    ' Lifecycle
    ' --------------------------------------------------

    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        MyBase.OnVisibleChanged(e)

        If Not Visible AndAlso _AnimationTimer IsNot Nothing Then
            SnapToTarget()
        End If
    End Sub

    Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
        If _AnimationTimer IsNot Nothing Then
            StopAnimation()
            _DisplayedProgress = Percentage / 100.0R
        End If

        MyBase.OnHandleDestroyed(e)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso _AnimationTimer IsNot Nothing Then
            RemoveHandler _AnimationTimer.Tick, AddressOf AnimationTimer_Tick

            StopAnimation()
            _AnimationTimer.Dispose()
        End If

        MyBase.Dispose(disposing)
    End Sub

End Class
