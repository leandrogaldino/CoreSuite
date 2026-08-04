Imports System.ComponentModel
Imports System.Threading
''' <summary>
''' Blocks an existing Windows Forms control with a customizable busy surface while one or more operations are active.
''' </summary>
''' <remarks>
''' Add the component to a form, assign <see cref="TargetControl"/>, and use <see cref="RunAsync(Func(Of CancellationToken, Task), CancellationToken)"/>, <see cref="BeginOperation"/>, or <see cref="ShowOverlay"/>. The target remains unchanged and the overlay is created only at run time.
''' </remarks>
<DefaultEvent("CancellationRequested")>
<DefaultProperty("TargetControl")>
<Description("Blocks an existing Windows Forms control with a customizable animated busy overlay.")>
<DesignerCategory("Component")>
<Designer(GetType(BusyOverlayDesigner))>
Partial Public Class BusyOverlay
    Inherits Component
    Private Const CategoryName As String = "BusyOverlay"
    Private ReadOnly _CancellationSources As New HashSet(Of CancellationTokenSource)
    Private _TargetControl As Control
    Private _ObservedScrollableParent As ScrollableControl
    Private _View As BusyOverlayView
    Private _Enabled As Boolean = True
    Private _ManualBusy As Boolean
    Private _OperationDisplayRequested As Boolean
    Private _OverlaySurfaceShown As Boolean
    Private _ActiveOperationCount As Integer
    Private _IsDisposed As Boolean
    Private _OverlayShownAt As DateTimeOffset?
    Private _MessageText As String = "Please wait..."
    Private _DetailText As String = String.Empty
    Private _MessageFont As Font = SystemFonts.MessageBoxFont
    Private _DetailFont As Font = SystemFonts.MessageBoxFont
    Private _MessageForeColor As Color = SystemColors.WindowText
    Private _DetailForeColor As Color = SystemColors.GrayText
    Private _OverlayColor As Color = SystemColors.Control
    Private _OverlayOpacity As Integer = 190
    Private _CaptureTarget As Boolean = True
    Private _ShowContentPanel As Boolean = True
    Private _ContentBackColor As Color = SystemColors.Window
    Private _ContentOpacity As Integer = 245
    Private _ContentBorderColor As Color = SystemColors.ControlDark
    Private _ContentBorderThickness As Integer
    Private _ContentCornerRadius As Integer = 8
    Private _ContentPadding As Integer = 20
    Private _ContentSpacing As Integer = 10
    Private _ContentMaximumWidth As Integer = 420
    Private _IndicatorStyle As BusyOverlayIndicatorStyle = BusyOverlayIndicatorStyle.Spinner
    Private _IndicatorColor As Color = SystemColors.Highlight
    Private _IndicatorTrackColor As Color = SystemColors.ControlDark
    Private _IndicatorSize As Integer = 32
    Private _IndicatorThickness As Integer = 4
    Private _AnimationInterval As Integer = 75
    Private _ProgressBarWidth As Integer = 220
    Private _ProgressBarHeight As Integer = 8
    Private _ProgressMinimum As Integer
    Private _ProgressMaximum As Integer = 100
    Private _ProgressValue As Integer
    Private _ShowProgressPercentage As Boolean = True
    Private _AllowCancellation As Boolean
    Private _CancelButtonText As String = "Cancel"
    Private _CancelButtonSize As New Size(90, 30)
    Private _UseWaitCursor As Boolean = True
    Private _BlockKeyboardInput As Boolean = True
    Private _OperationDisplayDelay As Integer = 150
    Private _MinimumOperationDisplayTime As Integer = 300
    Private _RestoreFocusControl As Control
    ''' <summary>
    ''' Initializes a new instance of the <see cref="BusyOverlay"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="BusyOverlay"/> class and adds it to the specified component container.
    ''' </summary>
    ''' <param name="container">The container that owns the component.</param>
    Public Sub New(container As IContainer)
        Me.New()
        ArgumentNullException.ThrowIfNull(container)
        container.Add(Me)
    End Sub
    ''' <summary>
    ''' Detaches the target, cancels active cancellable operations, and releases the run-time overlay surface.
    ''' </summary>
    ''' <param name="disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not _IsDisposed Then
            _IsDisposed = True
            For Each CancellationSource As CancellationTokenSource In _CancellationSources.ToArray()
                Try
                    CancellationSource.Cancel()
                Catch ex As ObjectDisposedException
                Catch ex As AggregateException
                End Try
            Next
            _CancellationSources.Clear()
            _ManualBusy = False
            _OperationDisplayRequested = False
            _ActiveOperationCount = 0
            DetachTargetControl()
            DisposeOverlayView()
        End If
        MyBase.Dispose(disposing)
    End Sub
    Private ReadOnly Property IsInDesignMode As Boolean
        Get
            Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse (Site IsNot Nothing AndAlso Site.DesignMode)
        End Get
    End Property
End Class
