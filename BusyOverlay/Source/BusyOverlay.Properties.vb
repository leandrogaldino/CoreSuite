Imports System.ComponentModel
Partial Public Class BusyOverlay
    ''' <summary>
    ''' Gets or sets the control whose client area is blocked while the component is busy.
    ''' </summary>
    ''' <value>A form or any child control, or <see langword="Nothing"/> when no target is assigned.</value>
    <Category(CategoryName)>
    <Description("Specifies the form or control covered while the component is busy.")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property TargetControl As Control
        Get
            Return _TargetControl
        End Get
        Set(value As Control)
            SetTargetControl(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the component can display its overlay.
    ''' </summary>
    ''' <value><see langword="True"/> to permit display; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the component can display its overlay.")>
    <DefaultValue(True)>
    Public Property Enabled As Boolean
        Get
            Return _Enabled
        End Get
        Set(value As Boolean)
            If _Enabled = value Then Return
            _Enabled = value
            SynchronizeOverlayVisibility()
        End Set
    End Property
    ''' <summary>
    ''' Gets a value indicating whether a manual, scoped, or asynchronous operation is active.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsBusy As Boolean
        Get
            Return _ManualBusy OrElse _ActiveOperationCount > 0
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the overlay surface is currently visible.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsOverlayVisible As Boolean
        Get
            Return _OverlaySurfaceShown AndAlso _View IsNot Nothing AndAlso Not _View.IsDisposed AndAlso _View.Visible
        End Get
    End Property
    ''' <summary>
    ''' Gets the number of active operations started by <see cref="BeginOperation"/> or a <c>RunAsync</c> overload.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property ActiveOperationCount As Integer
        Get
            Return _ActiveOperationCount
        End Get
    End Property
    ''' <summary>
    ''' Gets the number of active asynchronous operations that can receive cancellation.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property CancellableOperationCount As Integer
        Get
            Return _CancellationSources.Count
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether cancellation is currently available to the user.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property CanCancel As Boolean
        Get
            Return _AllowCancellation AndAlso _CancellationSources.Count > 0
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the primary message displayed in the center of the overlay.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the primary message displayed by the overlay.")>
    <DefaultValue("Please wait...")>
    Public Property MessageText As String
        Get
            Return _MessageText
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty)
            If String.Equals(_MessageText, NormalizedValue, StringComparison.Ordinal) Then Return
            _MessageText = NormalizedValue
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the optional secondary message displayed below the primary message.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies optional detail text displayed below the primary message.")>
    <DefaultValue("")>
    Public Property DetailText As String
        Get
            Return _DetailText
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty)
            If String.Equals(_DetailText, NormalizedValue, StringComparison.Ordinal) Then Return
            _DetailText = NormalizedValue
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the font used for the primary message.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the font used for the primary message.")>
    Public Property MessageFont As Font
        Get
            Return _MessageFont
        End Get
        Set(value As Font)
            ArgumentNullException.ThrowIfNull(value)
            If Equals(_MessageFont, value) Then Return
            _MessageFont = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the font used for the detail text and progress percentage.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the font used for detail text and progress percentage.")>
    Public Property DetailFont As Font
        Get
            Return _DetailFont
        End Get
        Set(value As Font)
            ArgumentNullException.ThrowIfNull(value)
            If Equals(_DetailFont, value) Then Return
            _DetailFont = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the primary message.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the color of the primary message.")>
    <DefaultValue(GetType(Color), "WindowText")>
    Public Property MessageForeColor As Color
        Get
            Return _MessageForeColor
        End Get
        Set(value As Color)
            SetColor(_MessageForeColor, value, NameOf(MessageForeColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the detail text and progress percentage.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the color of detail text and the progress percentage.")>
    <DefaultValue(GetType(Color), "GrayText")>
    Public Property DetailForeColor As Color
        Get
            Return _DetailForeColor
        End Get
        Set(value As Color)
            SetColor(_DetailForeColor, value, NameOf(DetailForeColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color applied across the target area.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the color applied across the blocked target area.")>
    <DefaultValue(GetType(Color), "Control")>
    Public Property OverlayColor As Color
        Get
            Return _OverlayColor
        End Get
        Set(value As Color)
            SetColor(_OverlayColor, value, NameOf(OverlayColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the opacity of <see cref="OverlayColor"/> from 0 through 255.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies overlay opacity from 0 through 255.")>
    <DefaultValue(190)>
    Public Property OverlayOpacity As Integer
        Get
            Return _OverlayOpacity
        End Get
        Set(value As Integer)
            SetRange(_OverlayOpacity, value, 0, 255, NameOf(OverlayOpacity))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the target is captured before the tint is drawn.
    ''' </summary>
    ''' <value><see langword="True"/> to preserve a visual snapshot beneath translucent overlays; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the target is captured beneath a translucent overlay.")>
    <DefaultValue(True)>
    Public Property CaptureTarget As Boolean
        Get
            Return _CaptureTarget
        End Get
        Set(value As Boolean)
            If _CaptureTarget = value Then Return
            _CaptureTarget = value
            If IsOverlayVisible Then CaptureTargetSnapshot()
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether messages and indicators are drawn on a centered content panel.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Determines whether overlay content is drawn on a centered panel.")>
    <DefaultValue(True)>
    Public Property ShowContentPanel As Boolean
        Get
            Return _ShowContentPanel
        End Get
        Set(value As Boolean)
            If _ShowContentPanel = value Then Return
            _ShowContentPanel = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the content panel background color.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the content panel background color.")>
    <DefaultValue(GetType(Color), "Window")>
    Public Property ContentBackColor As Color
        Get
            Return _ContentBackColor
        End Get
        Set(value As Color)
            SetColor(_ContentBackColor, value, NameOf(ContentBackColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the content panel opacity from 0 through 255.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies content panel opacity from 0 through 255.")>
    <DefaultValue(245)>
    Public Property ContentOpacity As Integer
        Get
            Return _ContentOpacity
        End Get
        Set(value As Integer)
            SetRange(_ContentOpacity, value, 0, 255, NameOf(ContentOpacity))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the content panel border color.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the content panel border color.")>
    <DefaultValue(GetType(Color), "ControlDark")>
    Public Property ContentBorderColor As Color
        Get
            Return _ContentBorderColor
        End Get
        Set(value As Color)
            SetColor(_ContentBorderColor, value, NameOf(ContentBorderColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the content panel border thickness from 0 through 10 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies content panel border thickness from 0 through 10 pixels.")>
    <DefaultValue(0)>
    Public Property ContentBorderThickness As Integer
        Get
            Return _ContentBorderThickness
        End Get
        Set(value As Integer)
            SetRange(_ContentBorderThickness, value, 0, 10, NameOf(ContentBorderThickness))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the content panel corner radius from 0 through 64 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies content panel corner radius from 0 through 64 pixels.")>
    <DefaultValue(8)>
    Public Property ContentCornerRadius As Integer
        Get
            Return _ContentCornerRadius
        End Get
        Set(value As Integer)
            SetRange(_ContentCornerRadius, value, 0, 64, NameOf(ContentCornerRadius))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the internal content panel padding from 0 through 64 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies content panel padding from 0 through 64 pixels.")>
    <DefaultValue(20)>
    Public Property ContentPadding As Integer
        Get
            Return _ContentPadding
        End Get
        Set(value As Integer)
            SetRange(_ContentPadding, value, 0, 64, NameOf(ContentPadding))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the space between visible content items from 0 through 64 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies spacing between visible content items from 0 through 64 pixels.")>
    <DefaultValue(10)>
    Public Property ContentSpacing As Integer
        Get
            Return _ContentSpacing
        End Get
        Set(value As Integer)
            SetRange(_ContentSpacing, value, 0, 64, NameOf(ContentSpacing))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the maximum content panel width from 120 through 2000 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies maximum content panel width from 120 through 2000 pixels.")>
    <DefaultValue(420)>
    Public Property ContentMaximumWidth As Integer
        Get
            Return _ContentMaximumWidth
        End Get
        Set(value As Integer)
            SetRange(_ContentMaximumWidth, value, 120, 2000, NameOf(ContentMaximumWidth))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the indicator drawn above the message.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the spinner, marquee, progress bar, or no indicator.")>
    <DefaultValue(BusyOverlayIndicatorStyle.Spinner)>
    Public Property IndicatorStyle As BusyOverlayIndicatorStyle
        Get
            Return _IndicatorStyle
        End Get
        Set(value As BusyOverlayIndicatorStyle)
            If Not [Enum].IsDefined(GetType(BusyOverlayIndicatorStyle), value) Then Throw New InvalidEnumArgumentException(NameOf(value), CInt(value), GetType(BusyOverlayIndicatorStyle))
            If _IndicatorStyle = value Then Return
            _IndicatorStyle = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the active indicator color.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the active spinner or progress color.")>
    <DefaultValue(GetType(Color), "Highlight")>
    Public Property IndicatorColor As Color
        Get
            Return _IndicatorColor
        End Get
        Set(value As Color)
            SetColor(_IndicatorColor, value, NameOf(IndicatorColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the inactive track color used by bar indicators.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the inactive track color used by bar indicators.")>
    <DefaultValue(GetType(Color), "ControlDark")>
    Public Property IndicatorTrackColor As Color
        Get
            Return _IndicatorTrackColor
        End Get
        Set(value As Color)
            SetColor(_IndicatorTrackColor, value, NameOf(IndicatorTrackColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the spinner diameter from 16 through 128 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies spinner diameter from 16 through 128 pixels.")>
    <DefaultValue(32)>
    Public Property IndicatorSize As Integer
        Get
            Return _IndicatorSize
        End Get
        Set(value As Integer)
            ValidateRange(value, 16, 128, NameOf(IndicatorSize))
            If _IndicatorThickness > value \ 2 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "IndicatorSize must be at least twice IndicatorThickness.")
            If _IndicatorSize = value Then Return
            _IndicatorSize = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the spinner line thickness from 1 through 32 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies spinner line thickness from 1 through 32 pixels.")>
    <DefaultValue(4)>
    Public Property IndicatorThickness As Integer
        Get
            Return _IndicatorThickness
        End Get
        Set(value As Integer)
            ValidateRange(value, 1, 32, NameOf(IndicatorThickness))
            If value > _IndicatorSize \ 2 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "IndicatorThickness cannot exceed half of IndicatorSize.")
            If _IndicatorThickness = value Then Return
            _IndicatorThickness = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the animation interval from 15 through 1000 milliseconds.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies animation interval from 15 through 1000 milliseconds.")>
    <DefaultValue(75)>
    Public Property AnimationInterval As Integer
        Get
            Return _AnimationInterval
        End Get
        Set(value As Integer)
            SetRange(_AnimationInterval, value, 15, 1000, NameOf(AnimationInterval))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the width of marquee and determinate progress bars from 60 through 1000 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies progress bar width from 60 through 1000 pixels.")>
    <DefaultValue(220)>
    Public Property ProgressBarWidth As Integer
        Get
            Return _ProgressBarWidth
        End Get
        Set(value As Integer)
            SetRange(_ProgressBarWidth, value, 60, 1000, NameOf(ProgressBarWidth))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the height of marquee and determinate progress bars from 2 through 64 pixels.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies progress bar height from 2 through 64 pixels.")>
    <DefaultValue(8)>
    Public Property ProgressBarHeight As Integer
        Get
            Return _ProgressBarHeight
        End Get
        Set(value As Integer)
            SetRange(_ProgressBarHeight, value, 2, 64, NameOf(ProgressBarHeight))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the lower bound of determinate progress.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the lower bound of determinate progress.")>
    <DefaultValue(0)>
    Public Property ProgressMinimum As Integer
        Get
            Return _ProgressMinimum
        End Get
        Set(value As Integer)
            If value >= _ProgressMaximum Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "ProgressMinimum must be less than ProgressMaximum.")
            If _ProgressMinimum = value Then Return
            _ProgressMinimum = value
            Dim progressAdjusted As Boolean = _ProgressValue < value
            If progressAdjusted Then _ProgressValue = value
            RefreshOverlayAppearance()
            If progressAdjusted Then OnProgressChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the upper bound of determinate progress.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the upper bound of determinate progress.")>
    <DefaultValue(100)>
    Public Property ProgressMaximum As Integer
        Get
            Return _ProgressMaximum
        End Get
        Set(value As Integer)
            If value <= _ProgressMinimum Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "ProgressMaximum must be greater than ProgressMinimum.")
            If _ProgressMaximum = value Then Return
            _ProgressMaximum = value
            Dim progressAdjusted As Boolean = _ProgressValue > value
            If progressAdjusted Then _ProgressValue = value
            RefreshOverlayAppearance()
            If progressAdjusted Then OnProgressChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the current determinate progress value.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the current determinate progress value.")>
    <DefaultValue(0)>
    Public Property ProgressValue As Integer
        Get
            Return _ProgressValue
        End Get
        Set(value As Integer)
            SetProgress(value, Nothing, True)
        End Set
    End Property
    ''' <summary>
    ''' Gets the normalized determinate progress percentage from 0 through 100.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property ProgressPercentage As Double
        Get
            Return (_ProgressValue - _ProgressMinimum) * 100.0R / (_ProgressMaximum - _ProgressMinimum)
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether a percentage is drawn below a determinate progress bar.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Determines whether determinate progress displays a percentage.")>
    <DefaultValue(True)>
    Public Property ShowProgressPercentage As Boolean
        Get
            Return _ShowProgressPercentage
        End Get
        Set(value As Boolean)
            If _ShowProgressPercentage = value Then Return
            _ShowProgressPercentage = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether cancellable operations display a cancel button.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Determines whether cancellable operations display a cancel button.")>
    <DefaultValue(False)>
    Public Property AllowCancellation As Boolean
        Get
            Return _AllowCancellation
        End Get
        Set(value As Boolean)
            If _AllowCancellation = value Then Return
            _AllowCancellation = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text displayed by the cancellation button.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the text displayed by the cancellation button.")>
    <DefaultValue("Cancel")>
    Public Property CancelButtonText As String
        Get
            Return _CancelButtonText
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty)
            If String.Equals(_CancelButtonText, NormalizedValue, StringComparison.Ordinal) Then Return
            _CancelButtonText = NormalizedValue
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the cancellation button size.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Specifies the cancellation button size.")>
    <DefaultValue(GetType(Size), "90, 30")>
    Public Property CancelButtonSize As Size
        Get
            Return _CancelButtonSize
        End Get
        Set(value As Size)
            If value.Width < 40 OrElse value.Height < 20 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "CancelButtonSize must be at least 40 by 20 pixels.")
            If _CancelButtonSize = value Then Return
            _CancelButtonSize = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the overlay uses the wait cursor.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Determines whether the overlay uses the wait cursor.")>
    <DefaultValue(True)>
    Public Property UseWaitCursor As Boolean
        Get
            Return _UseWaitCursor
        End Get
        Set(value As Boolean)
            If _UseWaitCursor = value Then Return
            _UseWaitCursor = value
            RefreshOverlayAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether keyboard focus moves to the overlay while it is visible.
    ''' </summary>
    ''' <value><see langword="True"/> to prevent the previously focused control from receiving keyboard input; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the overlay prevents its target from receiving keyboard input.")>
    <DefaultValue(True)>
    Public Property BlockKeyboardInput As Boolean
        Get
            Return _BlockKeyboardInput
        End Get
        Set(value As Boolean)
            If _BlockKeyboardInput = value Then Return
            _BlockKeyboardInput = value
            If IsOverlayVisible AndAlso value Then FocusOverlay()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how long a <c>RunAsync</c> operation may run before the overlay is displayed.
    ''' </summary>
    ''' <value>A delay from 0 through 60000 milliseconds. The default is 150 milliseconds.</value>
    <Category(CategoryName)>
    <Description("Specifies the RunAsync display delay from 0 through 60000 milliseconds.")>
    <DefaultValue(150)>
    Public Property OperationDisplayDelay As Integer
        Get
            Return _OperationDisplayDelay
        End Get
        Set(value As Integer)
            ValidateRange(value, 0, 60000, NameOf(OperationDisplayDelay))
            _OperationDisplayDelay = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how long an overlay displayed by <c>RunAsync</c> remains visible to avoid a brief flash.
    ''' </summary>
    ''' <value>A duration from 0 through 60000 milliseconds. The default is 300 milliseconds.</value>
    <Category(CategoryName)>
    <Description("Specifies the minimum RunAsync display time from 0 through 60000 milliseconds.")>
    <DefaultValue(300)>
    Public Property MinimumOperationDisplayTime As Integer
        Get
            Return _MinimumOperationDisplayTime
        End Get
        Set(value As Integer)
            ValidateRange(value, 0, 60000, NameOf(MinimumOperationDisplayTime))
            _MinimumOperationDisplayTime = value
        End Set
    End Property
    Private Sub SetColor(ByRef field As Color, value As Color, propertyName As String)
        If value.IsEmpty Then Throw New ArgumentException("The color cannot be empty.", propertyName)
        If field = value Then Return
        field = value
        RefreshOverlayAppearance()
    End Sub
    Private Sub SetRange(ByRef field As Integer, value As Integer, minimum As Integer, maximum As Integer, propertyName As String)
        ValidateRange(value, minimum, maximum, propertyName)
        If field = value Then Return
        field = value
        RefreshOverlayAppearance()
    End Sub
    Private Shared Sub ValidateRange(value As Integer, minimum As Integer, maximum As Integer, propertyName As String)
        If value < minimum OrElse value > maximum Then Throw New ArgumentOutOfRangeException(propertyName, value, $"The value must be from {minimum} through {maximum}.")
    End Sub
End Class
