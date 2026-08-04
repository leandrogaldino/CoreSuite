Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
''' <summary>
''' Specifies the position of the caption relative to the toggle switch.
''' </summary>
Public Enum ToggleButtonTextPosition
    ''' <summary>
    ''' Places the caption to the left of the switch.
    ''' </summary>
    Left
    ''' <summary>
    ''' Places the caption to the right of the switch.
    ''' </summary>
    Right
End Enum
''' <summary>
''' Specifies the easing function used by the toggle animation.
''' </summary>
Public Enum ToggleButtonAnimationEasing
    ''' <summary>
    ''' Uses constant-speed interpolation.
    ''' </summary>
    Linear
    ''' <summary>
    ''' Starts quickly and slows down near the destination.
    ''' </summary>
    EaseOutCubic
    ''' <summary>
    ''' Accelerates at the beginning and decelerates near the destination.
    ''' </summary>
    EaseInOutCubic
End Enum
''' <summary>
''' Defines the colors used to render one visual state of a <see cref="ToggleButton"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public NotInheritable Class ToggleButtonVisualStyle
    Private _TrackBackColor As Color
    Private _TrackBorderColor As Color
    Private _ThumbBackColor As Color
    Private _ThumbBorderColor As Color
    Private _CaptionColor As Color
    Private _StateTextColor As Color
    Private _ThumbGlyphColor As Color
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ToggleButtonVisualStyle"/> class.
    ''' </summary>
    Public Sub New()
        Me.New(Color.FromArgb(208, 213, 221), Color.FromArgb(184, 192, 204), Color.White, Color.FromArgb(226, 229, 234), Color.FromArgb(52, 64, 84), Color.FromArgb(102, 112, 133), Color.FromArgb(102, 112, 133))
    End Sub
    Friend Sub New(TrackBackColor As Color, TrackBorderColor As Color, ThumbBackColor As Color, ThumbBorderColor As Color, CaptionColor As Color, StateTextColor As Color, ThumbGlyphColor As Color)
        _TrackBackColor = TrackBackColor
        _TrackBorderColor = TrackBorderColor
        _ThumbBackColor = ThumbBackColor
        _ThumbBorderColor = ThumbBorderColor
        _CaptionColor = CaptionColor
        _StateTextColor = StateTextColor
        _ThumbGlyphColor = ThumbGlyphColor
    End Sub
    ''' <summary>
    ''' Gets or sets the track background color.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the track background color.")>
    Public Property TrackBackColor As Color
        Get
            Return _TrackBackColor
        End Get
        Set(Value As Color)
            If _TrackBackColor = Value Then Return
            _TrackBackColor = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the track border color.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the track border color.")>
    Public Property TrackBorderColor As Color
        Get
            Return _TrackBorderColor
        End Get
        Set(Value As Color)
            If _TrackBorderColor = Value Then Return
            _TrackBorderColor = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the thumb background color.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the thumb background color.")>
    Public Property ThumbBackColor As Color
        Get
            Return _ThumbBackColor
        End Get
        Set(Value As Color)
            If _ThumbBackColor = Value Then Return
            _ThumbBackColor = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the thumb border color.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the thumb border color.")>
    Public Property ThumbBorderColor As Color
        Get
            Return _ThumbBorderColor
        End Get
        Set(Value As Color)
            If _ThumbBorderColor = Value Then Return
            _ThumbBorderColor = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the external caption color.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the external caption color.")>
    Public Property CaptionColor As Color
        Get
            Return _CaptionColor
        End Get
        Set(Value As Color)
            If _CaptionColor = Value Then Return
            _CaptionColor = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the state text rendered inside the track.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the color of the state text rendered inside the track.")>
    Public Property StateTextColor As Color
        Get
            Return _StateTextColor
        End Get
        Set(Value As Color)
            If _StateTextColor = Value Then Return
            _StateTextColor = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the glyph rendered inside the thumb.
    ''' </summary>
    <NotifyParentProperty(True), Description("Gets or sets the color of the glyph rendered inside the thumb.")>
    Public Property ThumbGlyphColor As Color
        Get
            Return _ThumbGlyphColor
        End Get
        Set(Value As Color)
            If _ThumbGlyphColor = Value Then Return
            _ThumbGlyphColor = Value
            OnChanged()
        End Set
    End Property
    Friend Sub SetValues(TrackBackColor As Color, TrackBorderColor As Color, ThumbBackColor As Color, ThumbBorderColor As Color, CaptionColor As Color, StateTextColor As Color, ThumbGlyphColor As Color)
        _TrackBackColor = TrackBackColor
        _TrackBorderColor = TrackBorderColor
        _ThumbBackColor = ThumbBackColor
        _ThumbBorderColor = ThumbBorderColor
        _CaptionColor = CaptionColor
        _StateTextColor = StateTextColor
        _ThumbGlyphColor = ThumbGlyphColor
        OnChanged()
    End Sub
    Friend Function Matches(TrackBackColor As Color, TrackBorderColor As Color, ThumbBackColor As Color, ThumbBorderColor As Color, CaptionColor As Color, StateTextColor As Color, ThumbGlyphColor As Color) As Boolean
        Return _TrackBackColor = TrackBackColor AndAlso _TrackBorderColor = TrackBorderColor AndAlso _ThumbBackColor = ThumbBackColor AndAlso _ThumbBorderColor = ThumbBorderColor AndAlso _CaptionColor = CaptionColor AndAlso _StateTextColor = StateTextColor AndAlso _ThumbGlyphColor = ThumbGlyphColor
    End Function
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Returns a text representation of the visual style.
    ''' </summary>
    Public Overrides Function ToString() As String
        Return "Visual Style"
    End Function
End Class
''' <summary>
''' Represents a modern, animated and fully customizable two-state toggle button.
''' </summary>
<DefaultEvent("CheckedChanged"), DefaultProperty("Checked"), DefaultBindingProperty("Checked"), ToolboxItem(True)>
Public Class ToggleButton
    Inherits CheckBox
    Private Const DefaultDpi As Single = 96.0F
    Private ReadOnly _AnimationTimer As Timer
    Private ReadOnly _CheckedStyle As ToggleButtonVisualStyle
    Private ReadOnly _UncheckedStyle As ToggleButtonVisualStyle
    Private ReadOnly _DisabledStyle As ToggleButtonVisualStyle
    Private _TrackSize As Size = New Size(46, 26)
    Private _ThumbSize As Integer
    Private _ThumbPadding As Integer = 3
    Private _TrackCornerRadius As Integer = -1
    Private _TrackBorderThickness As Single = 1.0F
    Private _ThumbBorderThickness As Single = 1.0F
    Private _TextSpacing As Integer = 8
    Private _TextPosition As ToggleButtonTextPosition = ToggleButtonTextPosition.Right
    Private _ContentAlignment As ContentAlignment = ContentAlignment.MiddleLeft
    Private _MirrorInRightToLeft As Boolean = True
    Private _ShowStateText As Boolean
    Private _CheckedText As String = "ON"
    Private _UncheckedText As String = "OFF"
    Private _ShowThumbGlyph As Boolean
    Private _ShowThumbShadow As Boolean = True
    Private _ThumbShadowColor As Color = Color.FromArgb(55, 16, 24, 40)
    Private _ThumbShadowOffset As Point = New Point(0, 1)
    Private _FocusRingColor As Color = Color.FromArgb(105, 37, 99, 235)
    Private _FocusRingThickness As Single = 2.0F
    Private _FocusRingOffset As Integer = 1
    Private _HoverOverlayColor As Color = Color.White
    Private _HoverOverlayOpacity As Single = 0.12F
    Private _PressedOverlayColor As Color = Color.Black
    Private _PressedOverlayOpacity As Single = 0.1F
    Private _AnimationEnabled As Boolean = True
    Private _AnimationDuration As Integer = 160
    Private _AnimationEasing As ToggleButtonAnimationEasing = ToggleButtonAnimationEasing.EaseOutCubic
    Private _AnimationProgress As Single
    Private _AnimationStartProgress As Single
    Private _AnimationTargetProgress As Single
    Private _AnimationStartedAt As Long
    Private _MouseOver As Boolean
    Private _MousePressed As Boolean
    Private _KeyboardPressed As Boolean
    Private Structure ToggleLayout
        Public SwitchBounds As Rectangle
        Public TrackBounds As Rectangle
        Public CaptionBounds As Rectangle
    End Structure
    Private Structure StyleSnapshot
        Public TrackBackColor As Color
        Public TrackBorderColor As Color
        Public ThumbBackColor As Color
        Public ThumbBorderColor As Color
        Public CaptionColor As Color
        Public StateTextColor As Color
        Public ThumbGlyphColor As Color
    End Structure
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ToggleButton"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor Or ControlStyles.Selectable, True)
        _CheckedStyle = New ToggleButtonVisualStyle(Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), Color.White, Color.FromArgb(219, 234, 254), Color.FromArgb(31, 41, 55), Color.White, Color.FromArgb(37, 99, 235))
        _UncheckedStyle = New ToggleButtonVisualStyle(Color.FromArgb(208, 213, 221), Color.FromArgb(184, 192, 204), Color.White, Color.FromArgb(226, 229, 234), Color.FromArgb(52, 64, 84), Color.FromArgb(102, 112, 133), Color.FromArgb(102, 112, 133))
        _DisabledStyle = New ToggleButtonVisualStyle(Color.FromArgb(234, 236, 240), Color.FromArgb(208, 213, 221), Color.FromArgb(249, 250, 251), Color.FromArgb(234, 236, 240), Color.FromArgb(152, 162, 179), Color.FromArgb(152, 162, 179), Color.FromArgb(152, 162, 179))
        AddHandler _CheckedStyle.Changed, AddressOf VisualStyleChanged
        AddHandler _UncheckedStyle.Changed, AddressOf VisualStyleChanged
        AddHandler _DisabledStyle.Changed, AddressOf VisualStyleChanged
        _AnimationTimer = New Timer With {.Interval = 15}
        AddHandler _AnimationTimer.Tick, AddressOf AnimationTimerTick
        AutoSize = True
        BackColor = Color.Transparent
        Cursor = Cursors.Hand
        TabStop = True
        MyBase.Appearance = System.Windows.Forms.Appearance.Normal
        MyBase.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        MyBase.ThreeState = False
        _AnimationProgress = If(Checked, 1.0F, 0.0F)
    End Sub
    ''' <summary>
    ''' Gets the visual style used while the control is checked.
    ''' </summary>
    <Category("ToggleButton"), Description("Gets the visual style used while the control is checked."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property CheckedStyle As ToggleButtonVisualStyle
        Get
            Return _CheckedStyle
        End Get
    End Property
    ''' <summary>
    ''' Gets the visual style used while the control is unchecked.
    ''' </summary>
    <Category("ToggleButton"), Description("Gets the visual style used while the control is unchecked."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property UncheckedStyle As ToggleButtonVisualStyle
        Get
            Return _UncheckedStyle
        End Get
    End Property
    ''' <summary>
    ''' Gets the visual style used while the control is disabled.
    ''' </summary>
    <Category("ToggleButton"), Description("Gets the visual style used while the control is disabled."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property DisabledStyle As ToggleButtonVisualStyle
        Get
            Return _DisabledStyle
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the logical size of the toggle track.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Size), "46, 26"), Description("Gets or sets the logical size of the toggle track.")>
    Public Property TrackSize As Size
        Get
            Return _TrackSize
        End Get
        Set(Value As Size)
            Dim NewValue As New Size(Math.Max(18, Value.Width), Math.Max(12, Value.Height))
            If _TrackSize = NewValue Then Return
            _TrackSize = NewValue
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical thumb diameter, or zero to calculate it automatically.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(0), Description("Gets or sets the logical thumb diameter, or zero to calculate it automatically.")>
    Public Property ThumbSize As Integer
        Get
            Return _ThumbSize
        End Get
        Set(Value As Integer)
            Dim NewValue As Integer = Math.Max(0, Value)
            If _ThumbSize = NewValue Then Return
            _ThumbSize = NewValue
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical space between the thumb and the track edges.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(3), Description("Gets or sets the logical space between the thumb and the track edges.")>
    Public Property ThumbPadding As Integer
        Get
            Return _ThumbPadding
        End Get
        Set(Value As Integer)
            Dim NewValue As Integer = Math.Max(0, Value)
            If _ThumbPadding = NewValue Then Return
            _ThumbPadding = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical track corner radius, or minus one to use a pill shape.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(-1), Description("Gets or sets the logical track corner radius, or minus one to use a pill shape.")>
    Public Property TrackCornerRadius As Integer
        Get
            Return _TrackCornerRadius
        End Get
        Set(Value As Integer)
            Dim NewValue As Integer = Math.Max(-1, Value)
            If _TrackCornerRadius = NewValue Then Return
            _TrackCornerRadius = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical track border thickness.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Single), "1"), Description("Gets or sets the logical track border thickness.")>
    Public Property TrackBorderThickness As Single
        Get
            Return _TrackBorderThickness
        End Get
        Set(Value As Single)
            Dim NewValue As Single = Math.Max(0.0F, Value)
            If Math.Abs(_TrackBorderThickness - NewValue) < 0.001F Then Return
            _TrackBorderThickness = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical thumb border thickness.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Single), "1"), Description("Gets or sets the logical thumb border thickness.")>
    Public Property ThumbBorderThickness As Single
        Get
            Return _ThumbBorderThickness
        End Get
        Set(Value As Single)
            Dim NewValue As Single = Math.Max(0.0F, Value)
            If Math.Abs(_ThumbBorderThickness - NewValue) < 0.001F Then Return
            _ThumbBorderThickness = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical spacing between the switch and its caption.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(8), Description("Gets or sets the logical spacing between the switch and its caption.")>
    Public Property TextSpacing As Integer
        Get
            Return _TextSpacing
        End Get
        Set(Value As Integer)
            Dim NewValue As Integer = Math.Max(0, Value)
            If _TextSpacing = NewValue Then Return
            _TextSpacing = NewValue
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the caption position relative to the switch.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(ToggleButtonTextPosition), NameOf(ToggleButtonTextPosition.Right)), Description("Gets or sets the caption position relative to the switch.")>
    Public Property TextPosition As ToggleButtonTextPosition
        Get
            Return _TextPosition
        End Get
        Set(Value As ToggleButtonTextPosition)
            If _TextPosition = Value Then Return
            _TextPosition = Value
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the alignment of the complete switch and caption content.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(ContentAlignment), NameOf(ContentAlignment.MiddleLeft)), Description("Gets or sets the alignment of the complete switch and caption content.")>
    Public Property ToggleContentAlignment As ContentAlignment
        Get
            Return _ContentAlignment
        End Get
        Set(Value As ContentAlignment)
            If _ContentAlignment = Value Then Return
            _ContentAlignment = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether switch direction and caption position are mirrored in right-to-left layouts.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(True), Description("Gets or sets whether switch direction and caption position are mirrored in right-to-left layouts.")>
    Public Property MirrorInRightToLeft As Boolean
        Get
            Return _MirrorInRightToLeft
        End Get
        Set(Value As Boolean)
            If _MirrorInRightToLeft = Value Then Return
            _MirrorInRightToLeft = Value
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the current state text is rendered inside the track.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(False), Description("Gets or sets whether the current state text is rendered inside the track.")>
    Public Property ShowStateText As Boolean
        Get
            Return _ShowStateText
        End Get
        Set(Value As Boolean)
            If _ShowStateText = Value Then Return
            _ShowStateText = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text rendered inside the track while checked.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue("ON"), Description("Gets or sets the text rendered inside the track while checked.")>
    Public Property CheckedText As String
        Get
            Return _CheckedText
        End Get
        Set(Value As String)
            Dim NewValue As String = If(Value, String.Empty)
            If _CheckedText = NewValue Then Return
            _CheckedText = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text rendered inside the track while unchecked.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue("OFF"), Description("Gets or sets the text rendered inside the track while unchecked.")>
    Public Property UncheckedText As String
        Get
            Return _UncheckedText
        End Get
        Set(Value As String)
            Dim NewValue As String = If(Value, String.Empty)
            If _UncheckedText = NewValue Then Return
            _UncheckedText = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether a state glyph is rendered inside the thumb.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(False), Description("Gets or sets whether a state glyph is rendered inside the thumb.")>
    Public Property ShowThumbGlyph As Boolean
        Get
            Return _ShowThumbGlyph
        End Get
        Set(Value As Boolean)
            If _ShowThumbGlyph = Value Then Return
            _ShowThumbGlyph = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether a subtle shadow is rendered below the thumb.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(True), Description("Gets or sets whether a subtle shadow is rendered below the thumb.")>
    Public Property ShowThumbShadow As Boolean
        Get
            Return _ShowThumbShadow
        End Get
        Set(Value As Boolean)
            If _ShowThumbShadow = Value Then Return
            _ShowThumbShadow = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the thumb shadow color.
    ''' </summary>
    <Category("ToggleButton"), Description("Gets or sets the thumb shadow color.")>
    Public Property ThumbShadowColor As Color
        Get
            Return _ThumbShadowColor
        End Get
        Set(Value As Color)
            If _ThumbShadowColor = Value Then Return
            _ThumbShadowColor = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical thumb shadow offset.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Point), "0, 1"), Description("Gets or sets the logical thumb shadow offset.")>
    Public Property ThumbShadowOffset As Point
        Get
            Return _ThumbShadowOffset
        End Get
        Set(Value As Point)
            If _ThumbShadowOffset = Value Then Return
            _ThumbShadowOffset = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the keyboard focus ring color.
    ''' </summary>
    <Category("ToggleButton"), Description("Gets or sets the keyboard focus ring color.")>
    Public Property FocusRingColor As Color
        Get
            Return _FocusRingColor
        End Get
        Set(Value As Color)
            If _FocusRingColor = Value Then Return
            _FocusRingColor = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical keyboard focus ring thickness.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Single), "2"), Description("Gets or sets the logical keyboard focus ring thickness.")>
    Public Property FocusRingThickness As Single
        Get
            Return _FocusRingThickness
        End Get
        Set(Value As Single)
            Dim NewValue As Single = Math.Max(0.0F, Value)
            If Math.Abs(_FocusRingThickness - NewValue) < 0.001F Then Return
            _FocusRingThickness = NewValue
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical distance between the track and the keyboard focus ring.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(1), Description("Gets or sets the logical distance between the track and the keyboard focus ring.")>
    Public Property FocusRingOffset As Integer
        Get
            Return _FocusRingOffset
        End Get
        Set(Value As Integer)
            Dim NewValue As Integer = Math.Max(0, Value)
            If _FocusRingOffset = NewValue Then Return
            _FocusRingOffset = NewValue
            NotifyLayoutChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color blended over the control while the pointer is hovering.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Color), "White"), Description("Gets or sets the color blended over the control while the pointer is hovering.")>
    Public Property HoverOverlayColor As Color
        Get
            Return _HoverOverlayColor
        End Get
        Set(Value As Color)
            If _HoverOverlayColor = Value Then Return
            _HoverOverlayColor = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the hover overlay opacity from zero to one.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Single), "0.12"), Description("Gets or sets the hover overlay opacity from zero to one.")>
    Public Property HoverOverlayOpacity As Single
        Get
            Return _HoverOverlayOpacity
        End Get
        Set(Value As Single)
            Dim NewValue As Single = Clamp(Value, 0.0F, 1.0F)
            If Math.Abs(_HoverOverlayOpacity - NewValue) < 0.001F Then Return
            _HoverOverlayOpacity = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color blended over the control while it is pressed.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Color), "Black"), Description("Gets or sets the color blended over the control while it is pressed.")>
    Public Property PressedOverlayColor As Color
        Get
            Return _PressedOverlayColor
        End Get
        Set(Value As Color)
            If _PressedOverlayColor = Value Then Return
            _PressedOverlayColor = Value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the pressed overlay opacity from zero to one.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(Single), "0.1"), Description("Gets or sets the pressed overlay opacity from zero to one.")>
    Public Property PressedOverlayOpacity As Single
        Get
            Return _PressedOverlayOpacity
        End Get
        Set(Value As Single)
            Dim NewValue As Single = Clamp(Value, 0.0F, 1.0F)
            If Math.Abs(_PressedOverlayOpacity - NewValue) < 0.001F Then Return
            _PressedOverlayOpacity = NewValue
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether checked-state transitions are animated.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(True), Description("Gets or sets whether checked-state transitions are animated.")>
    Public Property AnimationEnabled As Boolean
        Get
            Return _AnimationEnabled
        End Get
        Set(Value As Boolean)
            If _AnimationEnabled = Value Then Return
            _AnimationEnabled = Value
            If Not Value Then SnapAnimationToState()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the animation duration in milliseconds.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(160), Description("Gets or sets the animation duration in milliseconds.")>
    Public Property AnimationDuration As Integer
        Get
            Return _AnimationDuration
        End Get
        Set(Value As Integer)
            Dim NewValue As Integer = Math.Max(0, Math.Min(2000, Value))
            If _AnimationDuration = NewValue Then Return
            _AnimationDuration = NewValue
            If NewValue = 0 Then SnapAnimationToState()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the easing function used by the checked-state animation.
    ''' </summary>
    <Category("ToggleButton"), DefaultValue(GetType(ToggleButtonAnimationEasing), NameOf(ToggleButtonAnimationEasing.EaseOutCubic)), Description("Gets or sets the easing function used by the checked-state animation.")>
    Public Property AnimationEasing As ToggleButtonAnimationEasing
        Get
            Return _AnimationEasing
        End Get
        Set(Value As ToggleButtonAnimationEasing)
            _AnimationEasing = Value
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property Appearance As System.Windows.Forms.Appearance
        Get
            Return System.Windows.Forms.Appearance.Normal
        End Get
        Set(Value As System.Windows.Forms.Appearance)
            MyBase.Appearance = System.Windows.Forms.Appearance.Normal
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property ThreeState As Boolean
        Get
            Return False
        End Get
        Set(Value As Boolean)
            MyBase.ThreeState = False
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property FlatStyle As System.Windows.Forms.FlatStyle
        Get
            Return System.Windows.Forms.FlatStyle.Flat
        End Get
        Set(Value As System.Windows.Forms.FlatStyle)
            MyBase.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property Image As Image
        Get
            Return Nothing
        End Get
        Set(Value As Image)
            MyBase.Image = Nothing
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property ImageAlign As ContentAlignment
        Get
            Return ContentAlignment.MiddleCenter
        End Get
        Set(Value As ContentAlignment)
            MyBase.ImageAlign = ContentAlignment.MiddleCenter
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property TextImageRelation As System.Windows.Forms.TextImageRelation
        Get
            Return System.Windows.Forms.TextImageRelation.Overlay
        End Get
        Set(Value As System.Windows.Forms.TextImageRelation)
            MyBase.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property CheckAlign As ContentAlignment
        Get
            Return ContentAlignment.MiddleLeft
        End Get
        Set(Value As ContentAlignment)
            MyBase.CheckAlign = ContentAlignment.MiddleLeft
        End Set
    End Property
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property UseVisualStyleBackColor As Boolean
        Get
            Return False
        End Get
        Set(Value As Boolean)
            MyBase.UseVisualStyleBackColor = False
        End Set
    End Property
    ''' <summary>
    ''' Calculates the preferred size of the control.
    ''' </summary>
    Public Overrides Function GetPreferredSize(ProposedSize As Size) As Size
        Dim FocusReserve As Integer = GetFocusReserve()
        Dim SwitchWidth As Integer = ScaleValue(_TrackSize.Width) + FocusReserve * 2
        Dim SwitchHeight As Integer = ScaleValue(_TrackSize.Height) + FocusReserve * 2
        Dim CaptionSize As Size = MeasureCaption()
        Dim Spacing As Integer = If(String.IsNullOrEmpty(Text), 0, ScaleValue(_TextSpacing))
        Dim PreferredWidth As Integer = SwitchWidth + Spacing + CaptionSize.Width + Padding.Horizontal
        Dim PreferredHeight As Integer = Math.Max(SwitchHeight, CaptionSize.Height) + Padding.Vertical
        Return New Size(Math.Max(1, PreferredWidth), Math.Max(1, PreferredHeight))
    End Function
    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return New Size(96, 32)
        End Get
    End Property
    Protected Overrides Sub OnHandleCreated(E As EventArgs)
        MyBase.OnHandleCreated(E)
        SnapAnimationToState()
    End Sub
    Protected Overrides Sub OnCheckedChanged(E As EventArgs)
        MyBase.OnCheckedChanged(E)
        StartStateAnimation()
    End Sub
    Protected Overrides Sub OnTextChanged(E As EventArgs)
        MyBase.OnTextChanged(E)
        NotifyLayoutChanged()
    End Sub
    Protected Overrides Sub OnFontChanged(E As EventArgs)
        MyBase.OnFontChanged(E)
        NotifyLayoutChanged()
    End Sub
    Protected Overrides Sub OnPaddingChanged(E As EventArgs)
        MyBase.OnPaddingChanged(E)
        NotifyLayoutChanged()
    End Sub
    Protected Overrides Sub OnRightToLeftChanged(E As EventArgs)
        MyBase.OnRightToLeftChanged(E)
        NotifyLayoutChanged()
    End Sub
    Protected Overrides Sub OnEnabledChanged(E As EventArgs)
        MyBase.OnEnabledChanged(E)
        If Not Enabled Then
            _MousePressed = False
            _KeyboardPressed = False
        End If
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseEnter(E As EventArgs)
        MyBase.OnMouseEnter(E)
        _MouseOver = True
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseLeave(E As EventArgs)
        MyBase.OnMouseLeave(E)
        _MouseOver = False
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseDown(E As MouseEventArgs)
        If E.Button = MouseButtons.Left AndAlso Enabled Then
            _MousePressed = True
            Invalidate()
        End If
        MyBase.OnMouseDown(E)
    End Sub
    Protected Overrides Sub OnMouseUp(E As MouseEventArgs)
        MyBase.OnMouseUp(E)
        If _MousePressed Then
            _MousePressed = False
            Invalidate()
        End If
    End Sub
    Protected Overrides Sub OnKeyDown(E As KeyEventArgs)
        If E.KeyCode = Keys.Space AndAlso Enabled AndAlso Not _KeyboardPressed Then
            _KeyboardPressed = True
            Invalidate()
        End If
        MyBase.OnKeyDown(E)
    End Sub
    Protected Overrides Sub OnKeyUp(E As KeyEventArgs)
        MyBase.OnKeyUp(E)
        If E.KeyCode = Keys.Space AndAlso _KeyboardPressed Then
            _KeyboardPressed = False
            Invalidate()
        End If
    End Sub
    Protected Overrides Sub OnGotFocus(E As EventArgs)
        MyBase.OnGotFocus(E)
        Invalidate()
    End Sub
    Protected Overrides Sub OnLostFocus(E As EventArgs)
        MyBase.OnLostFocus(E)
        _KeyboardPressed = False
        Invalidate()
    End Sub
    Protected Overrides Sub OnPaintBackground(E As PaintEventArgs)
        If BackColor = Color.Transparent AndAlso Parent IsNot Nothing Then
            MyBase.OnPaintBackground(E)
        Else
            E.Graphics.Clear(BackColor)
        End If
    End Sub
    Protected Overrides Sub OnPaint(E As PaintEventArgs)
        Dim Graphics As Graphics = E.Graphics
        Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
        Dim Layout As ToggleLayout = CalculateLayout()
        If Layout.TrackBounds.Width <= 0 OrElse Layout.TrackBounds.Height <= 0 Then Return
        Dim Style As StyleSnapshot = GetCurrentStyle()
        Dim TrackRectangle As New RectangleF(Layout.TrackBounds.X, Layout.TrackBounds.Y, Layout.TrackBounds.Width, Layout.TrackBounds.Height)
        Dim TrackRadius As Single = GetTrackRadius(TrackRectangle)
        DrawFocusRing(Graphics, TrackRectangle, TrackRadius)
        Using TrackPath As GraphicsPath = CreateRoundedRectangle(TrackRectangle, TrackRadius)
            Using TrackBrush As New SolidBrush(Style.TrackBackColor)
                Graphics.FillPath(TrackBrush, TrackPath)
            End Using
            Dim TrackBorderWidth As Single = ScaleSingle(_TrackBorderThickness)
            If TrackBorderWidth > 0.0F AndAlso Style.TrackBorderColor.A > 0 Then
                Using TrackPen As New Pen(Style.TrackBorderColor, TrackBorderWidth) With {.Alignment = PenAlignment.Inset}
                    Graphics.DrawPath(TrackPen, TrackPath)
                End Using
            End If
        End Using
        Dim ThumbRectangle As RectangleF = GetThumbRectangle(TrackRectangle)
        If _ShowStateText Then DrawStateText(Graphics, TrackRectangle, ThumbRectangle, Style.StateTextColor)
        If _ShowThumbShadow Then DrawThumbShadow(Graphics, ThumbRectangle)
        Using ThumbBrush As New SolidBrush(Style.ThumbBackColor)
            Graphics.FillEllipse(ThumbBrush, ThumbRectangle)
        End Using
        Dim ThumbBorderWidth As Single = ScaleSingle(_ThumbBorderThickness)
        If ThumbBorderWidth > 0.0F AndAlso Style.ThumbBorderColor.A > 0 Then
            Using ThumbPen As New Pen(Style.ThumbBorderColor, ThumbBorderWidth) With {.Alignment = PenAlignment.Inset}
                Graphics.DrawEllipse(ThumbPen, ThumbRectangle)
            End Using
        End If
        If _ShowThumbGlyph Then DrawThumbGlyph(Graphics, ThumbRectangle, Style.ThumbGlyphColor)
        DrawCaption(Graphics, Layout.CaptionBounds, Style.CaptionColor)
    End Sub
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            RemoveHandler _AnimationTimer.Tick, AddressOf AnimationTimerTick
            _AnimationTimer.Dispose()
            RemoveHandler _CheckedStyle.Changed, AddressOf VisualStyleChanged
            RemoveHandler _UncheckedStyle.Changed, AddressOf VisualStyleChanged
            RemoveHandler _DisabledStyle.Changed, AddressOf VisualStyleChanged
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Private Sub VisualStyleChanged(Sender As Object, E As EventArgs)
        Invalidate()
    End Sub
    Private Sub NotifyLayoutChanged()
        PerformLayout()
        Invalidate()
        If AutoSize Then
            Dim Preferred As Size = GetPreferredSize(Size.Empty)
            If Size <> Preferred Then Size = Preferred
        End If
    End Sub
    Private Function CalculateLayout() As ToggleLayout
        Dim Result As New ToggleLayout
        Dim Content As New Rectangle(Padding.Left, Padding.Top, Math.Max(0, ClientSize.Width - Padding.Horizontal), Math.Max(0, ClientSize.Height - Padding.Vertical))
        Dim FocusReserve As Integer = GetFocusReserve()
        Dim SwitchSize As New Size(ScaleValue(_TrackSize.Width) + FocusReserve * 2, ScaleValue(_TrackSize.Height) + FocusReserve * 2)
        Dim CaptionSize As Size = MeasureCaption()
        Dim HasCaption As Boolean = Not String.IsNullOrEmpty(Text)
        Dim Spacing As Integer = If(HasCaption, ScaleValue(_TextSpacing), 0)
        Dim GroupSize As New Size(SwitchSize.Width + Spacing + CaptionSize.Width, Math.Max(SwitchSize.Height, CaptionSize.Height))
        GroupSize.Width = Math.Min(GroupSize.Width, Content.Width)
        GroupSize.Height = Math.Min(GroupSize.Height, Content.Height)
        Dim GroupBounds As Rectangle = AlignRectangle(Content, GroupSize, _ContentAlignment)
        Dim EffectivePosition As ToggleButtonTextPosition = _TextPosition
        If _MirrorInRightToLeft AndAlso RightToLeft = RightToLeft.Yes Then EffectivePosition = If(EffectivePosition = ToggleButtonTextPosition.Left, ToggleButtonTextPosition.Right, ToggleButtonTextPosition.Left)
        If Not HasCaption Then
            Result.SwitchBounds = New Rectangle(GroupBounds.Left + Math.Max(0, (GroupBounds.Width - SwitchSize.Width) \ 2), GroupBounds.Top + Math.Max(0, (GroupBounds.Height - SwitchSize.Height) \ 2), Math.Min(SwitchSize.Width, GroupBounds.Width), Math.Min(SwitchSize.Height, GroupBounds.Height))
            Result.CaptionBounds = Rectangle.Empty
        ElseIf EffectivePosition = ToggleButtonTextPosition.Right Then
            Result.SwitchBounds = New Rectangle(GroupBounds.Left, GroupBounds.Top + Math.Max(0, (GroupBounds.Height - SwitchSize.Height) \ 2), Math.Min(SwitchSize.Width, GroupBounds.Width), Math.Min(SwitchSize.Height, GroupBounds.Height))
            Dim CaptionLeft As Integer = Math.Min(GroupBounds.Right, Result.SwitchBounds.Right + Spacing)
            Result.CaptionBounds = New Rectangle(CaptionLeft, GroupBounds.Top, Math.Max(0, GroupBounds.Right - CaptionLeft), GroupBounds.Height)
        Else
            Dim SwitchLeft As Integer = Math.Max(GroupBounds.Left, GroupBounds.Right - SwitchSize.Width)
            Result.SwitchBounds = New Rectangle(SwitchLeft, GroupBounds.Top + Math.Max(0, (GroupBounds.Height - SwitchSize.Height) \ 2), Math.Min(SwitchSize.Width, GroupBounds.Width), Math.Min(SwitchSize.Height, GroupBounds.Height))
            Result.CaptionBounds = New Rectangle(GroupBounds.Left, GroupBounds.Top, Math.Max(0, SwitchLeft - Spacing - GroupBounds.Left), GroupBounds.Height)
        End If
        Dim TrackWidth As Integer = Math.Min(ScaleValue(_TrackSize.Width), Math.Max(0, Result.SwitchBounds.Width - FocusReserve * 2))
        Dim TrackHeight As Integer = Math.Min(ScaleValue(_TrackSize.Height), Math.Max(0, Result.SwitchBounds.Height - FocusReserve * 2))
        Result.TrackBounds = New Rectangle(Result.SwitchBounds.Left + Math.Max(0, (Result.SwitchBounds.Width - TrackWidth) \ 2), Result.SwitchBounds.Top + Math.Max(0, (Result.SwitchBounds.Height - TrackHeight) \ 2), TrackWidth, TrackHeight)
        Return Result
    End Function
    Private Function MeasureCaption() As Size
        If String.IsNullOrEmpty(Text) Then Return Size.Empty
        Dim Flags As TextFormatFlags = TextFormatFlags.NoPadding Or TextFormatFlags.SingleLine
        If Not UseMnemonic Then Flags = Flags Or TextFormatFlags.NoPrefix
        Return TextRenderer.MeasureText(Text, Font, New Size(Integer.MaxValue, Integer.MaxValue), Flags)
    End Function
    Private Function GetFocusReserve() As Integer
        Return ScaleValue(_FocusRingOffset) + CInt(Math.Ceiling(ScaleSingle(_FocusRingThickness)))
    End Function
    Private Function GetCurrentStyle() As StyleSnapshot
        Dim Source As ToggleButtonVisualStyle = If(Enabled, If(Checked, _CheckedStyle, _UncheckedStyle), _DisabledStyle)
        Dim Result As New StyleSnapshot With {
            .TrackBackColor = Source.TrackBackColor,
            .TrackBorderColor = Source.TrackBorderColor,
            .ThumbBackColor = Source.ThumbBackColor,
            .ThumbBorderColor = Source.ThumbBorderColor,
            .CaptionColor = Source.CaptionColor,
            .StateTextColor = Source.StateTextColor,
            .ThumbGlyphColor = Source.ThumbGlyphColor
        }
        If Enabled Then
            If _MousePressed OrElse _KeyboardPressed Then
                ApplyOverlay(Result, _PressedOverlayColor, _PressedOverlayOpacity)
            ElseIf _MouseOver Then
                ApplyOverlay(Result, _HoverOverlayColor, _HoverOverlayOpacity)
            End If
        End If
        Return Result
    End Function
    Private Shared Sub ApplyOverlay(ByRef Style As StyleSnapshot, OverlayColor As Color, Opacity As Single)
        Style.TrackBackColor = BlendColor(Style.TrackBackColor, OverlayColor, Opacity)
        Style.TrackBorderColor = BlendColor(Style.TrackBorderColor, OverlayColor, Opacity)
        Style.ThumbBackColor = BlendColor(Style.ThumbBackColor, OverlayColor, Opacity * 0.65F)
        Style.ThumbBorderColor = BlendColor(Style.ThumbBorderColor, OverlayColor, Opacity * 0.65F)
    End Sub
    Private Function GetThumbRectangle(TrackRectangle As RectangleF) As RectangleF
        Dim PaddingValue As Single = ScaleValue(_ThumbPadding)
        Dim AvailableHeight As Single = Math.Max(1.0F, TrackRectangle.Height - PaddingValue * 2.0F)
        Dim RequestedSize As Single = If(_ThumbSize > 0, ScaleValue(_ThumbSize), AvailableHeight)
        Dim Diameter As Single = Math.Max(1.0F, Math.Min(RequestedSize, AvailableHeight))
        Dim InnerLeft As Single = TrackRectangle.Left + PaddingValue
        Dim InnerRight As Single = TrackRectangle.Right - PaddingValue
        Dim Travel As Single = Math.Max(0.0F, InnerRight - InnerLeft - Diameter)
        Dim VisualProgress As Single = _AnimationProgress
        If _MirrorInRightToLeft AndAlso RightToLeft = RightToLeft.Yes Then VisualProgress = 1.0F - VisualProgress
        Dim X As Single = InnerLeft + Travel * Clamp(VisualProgress, 0.0F, 1.0F)
        Dim Y As Single = TrackRectangle.Top + (TrackRectangle.Height - Diameter) / 2.0F
        Return New RectangleF(X, Y, Diameter, Diameter)
    End Function
    Private Function GetTrackRadius(TrackRectangle As RectangleF) As Single
        If _TrackCornerRadius < 0 Then Return TrackRectangle.Height / 2.0F
        Return Math.Min(ScaleValue(_TrackCornerRadius), Math.Min(TrackRectangle.Width, TrackRectangle.Height) / 2.0F)
    End Function
    Private Sub DrawFocusRing(Graphics As Graphics, TrackRectangle As RectangleF, TrackRadius As Single)
        If Not Focused OrElse Not ShowFocusCues OrElse _FocusRingThickness <= 0.0F OrElse _FocusRingColor.A = 0 Then Return
        Dim Offset As Single = ScaleValue(_FocusRingOffset)
        Dim FocusRectangle As RectangleF = TrackRectangle
        FocusRectangle.Inflate(Offset, Offset)
        Using FocusPath As GraphicsPath = CreateRoundedRectangle(FocusRectangle, TrackRadius + Offset)
            Using FocusPen As New Pen(_FocusRingColor, ScaleSingle(_FocusRingThickness))
                Graphics.DrawPath(FocusPen, FocusPath)
            End Using
        End Using
    End Sub
    Private Sub DrawStateText(Graphics As Graphics, TrackRectangle As RectangleF, ThumbRectangle As RectangleF, TextColor As Color)
        Dim StateValue As String = If(Checked, _CheckedText, _UncheckedText)
        If String.IsNullOrEmpty(StateValue) Then Return
        Dim HorizontalInset As Integer = Math.Max(1, ScaleValue(3))
        Dim TextBounds As Rectangle
        If ThumbRectangle.Left + ThumbRectangle.Width / 2.0F >= TrackRectangle.Left + TrackRectangle.Width / 2.0F Then
            TextBounds = Rectangle.FromLTRB(CInt(Math.Ceiling(TrackRectangle.Left)) + HorizontalInset, CInt(Math.Ceiling(TrackRectangle.Top)), CInt(Math.Floor(ThumbRectangle.Left)) - HorizontalInset, CInt(Math.Floor(TrackRectangle.Bottom)))
        Else
            TextBounds = Rectangle.FromLTRB(CInt(Math.Ceiling(ThumbRectangle.Right)) + HorizontalInset, CInt(Math.Ceiling(TrackRectangle.Top)), CInt(Math.Floor(TrackRectangle.Right)) - HorizontalInset, CInt(Math.Floor(TrackRectangle.Bottom)))
        End If
        If TextBounds.Width <= 3 OrElse TextBounds.Height <= 3 Then Return
        Dim StateFontSize As Single = Math.Max(6.0F, Math.Min(Font.SizeInPoints, TrackRectangle.Height * 0.28F))
        Using StateFont As New Font(Font.FontFamily, StateFontSize, FontStyle.Bold, GraphicsUnit.Point)
            Dim Flags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.SingleLine Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPadding Or TextFormatFlags.NoPrefix
            TextRenderer.DrawText(Graphics, StateValue, StateFont, TextBounds, TextColor, Color.Transparent, Flags)
        End Using
    End Sub
    Private Sub DrawThumbShadow(Graphics As Graphics, ThumbRectangle As RectangleF)
        If _ThumbShadowColor.A = 0 Then Return
        Dim ShadowRectangle As RectangleF = ThumbRectangle
        ShadowRectangle.Offset(ScaleValue(_ThumbShadowOffset.X), ScaleValue(_ThumbShadowOffset.Y))
        ShadowRectangle.Inflate(ScaleSingle(0.35F), ScaleSingle(0.35F))
        Using ShadowBrush As New SolidBrush(_ThumbShadowColor)
            Graphics.FillEllipse(ShadowBrush, ShadowRectangle)
        End Using
    End Sub
    Private Sub DrawThumbGlyph(Graphics As Graphics, ThumbRectangle As RectangleF, GlyphColor As Color)
        If GlyphColor.A = 0 OrElse ThumbRectangle.Width < ScaleValue(10) Then Return
        Dim PenWidth As Single = Math.Max(1.2F, ThumbRectangle.Width * 0.09F)
        Using GlyphPen As New Pen(GlyphColor, PenWidth) With {.StartCap = LineCap.Round, .EndCap = LineCap.Round, .LineJoin = LineJoin.Round}
            If Checked Then
                Dim Point1 As New PointF(ThumbRectangle.Left + ThumbRectangle.Width * 0.27F, ThumbRectangle.Top + ThumbRectangle.Height * 0.52F)
                Dim Point2 As New PointF(ThumbRectangle.Left + ThumbRectangle.Width * 0.43F, ThumbRectangle.Top + ThumbRectangle.Height * 0.68F)
                Dim Point3 As New PointF(ThumbRectangle.Left + ThumbRectangle.Width * 0.74F, ThumbRectangle.Top + ThumbRectangle.Height * 0.34F)
                Graphics.DrawLines(GlyphPen, {Point1, Point2, Point3})
            Else
                Dim Y As Single = ThumbRectangle.Top + ThumbRectangle.Height * 0.5F
                Graphics.DrawLine(GlyphPen, ThumbRectangle.Left + ThumbRectangle.Width * 0.33F, Y, ThumbRectangle.Left + ThumbRectangle.Width * 0.67F, Y)
            End If
        End Using
    End Sub
    Private Sub DrawCaption(Graphics As Graphics, CaptionBounds As Rectangle, CaptionColor As Color)
        If String.IsNullOrEmpty(Text) OrElse CaptionBounds.Width <= 0 OrElse CaptionBounds.Height <= 0 Then Return
        Dim Flags As TextFormatFlags = GetTextFlags(TextAlign)
        TextRenderer.DrawText(Graphics, Text, Font, CaptionBounds, CaptionColor, Color.Transparent, Flags)
    End Sub
    Private Function GetTextFlags(Alignment As ContentAlignment) As TextFormatFlags
        Dim Flags As TextFormatFlags = TextFormatFlags.SingleLine Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPadding
        If Not UseMnemonic Then Flags = Flags Or TextFormatFlags.NoPrefix
        If RightToLeft = RightToLeft.Yes Then Flags = Flags Or TextFormatFlags.RightToLeft
        Select Case Alignment
            Case ContentAlignment.TopLeft, ContentAlignment.MiddleLeft, ContentAlignment.BottomLeft
                Flags = Flags Or TextFormatFlags.Left
            Case ContentAlignment.TopCenter, ContentAlignment.MiddleCenter, ContentAlignment.BottomCenter
                Flags = Flags Or TextFormatFlags.HorizontalCenter
            Case Else
                Flags = Flags Or TextFormatFlags.Right
        End Select
        Select Case Alignment
            Case ContentAlignment.TopLeft, ContentAlignment.TopCenter, ContentAlignment.TopRight
                Flags = Flags Or TextFormatFlags.Top
            Case ContentAlignment.MiddleLeft, ContentAlignment.MiddleCenter, ContentAlignment.MiddleRight
                Flags = Flags Or TextFormatFlags.VerticalCenter
            Case Else
                Flags = Flags Or TextFormatFlags.Bottom
        End Select
        Return Flags
    End Function
    Private Sub StartStateAnimation()
        Dim Target As Single = If(Checked, 1.0F, 0.0F)
        If Not _AnimationEnabled OrElse _AnimationDuration <= 0 OrElse IsInDesignMode OrElse Not IsHandleCreated Then
            _AnimationTimer.Stop()
            _AnimationProgress = Target
            Invalidate()
            Return
        End If
        _AnimationStartProgress = _AnimationProgress
        _AnimationTargetProgress = Target
        _AnimationStartedAt = Environment.TickCount64
        _AnimationTimer.Start()
        Invalidate()
    End Sub
    Private Sub AnimationTimerTick(Sender As Object, E As EventArgs)
        Dim Elapsed As Long = Math.Max(0, Environment.TickCount64 - _AnimationStartedAt)
        Dim RawProgress As Single = Clamp(CSng(Elapsed / CDbl(Math.Max(1, _AnimationDuration))), 0.0F, 1.0F)
        Dim EasedProgress As Single = ApplyEasing(RawProgress, _AnimationEasing)
        _AnimationProgress = _AnimationStartProgress + (_AnimationTargetProgress - _AnimationStartProgress) * EasedProgress
        If RawProgress >= 1.0F Then
            _AnimationProgress = _AnimationTargetProgress
            _AnimationTimer.Stop()
        End If
        Invalidate()
    End Sub
    Private Sub SnapAnimationToState()
        _AnimationTimer.Stop()
        _AnimationProgress = If(Checked, 1.0F, 0.0F)
        _AnimationTargetProgress = _AnimationProgress
        Invalidate()
    End Sub
    Private ReadOnly Property IsInDesignMode As Boolean
        Get
            Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse DesignMode
        End Get
    End Property
    Private Function ScaleValue(Value As Integer) As Integer
        Return CInt(Math.Round(Value * GetDpiScale()))
    End Function
    Private Function ScaleSingle(Value As Single) As Single
        Return Value * GetDpiScale()
    End Function
    Private Function GetDpiScale() As Single
        Dim CurrentDpi As Integer = If(IsHandleCreated, DeviceDpi, CInt(DefaultDpi))
        Return Math.Max(1.0F, CurrentDpi / DefaultDpi)
    End Function
    Private Shared Function AlignRectangle(Container As Rectangle, Size As Size, Alignment As ContentAlignment) As Rectangle
        Dim X As Integer
        Dim Y As Integer
        Select Case Alignment
            Case ContentAlignment.TopCenter, ContentAlignment.MiddleCenter, ContentAlignment.BottomCenter
                X = Container.Left + (Container.Width - Size.Width) \ 2
            Case ContentAlignment.TopRight, ContentAlignment.MiddleRight, ContentAlignment.BottomRight
                X = Container.Right - Size.Width
            Case Else
                X = Container.Left
        End Select
        Select Case Alignment
            Case ContentAlignment.MiddleLeft, ContentAlignment.MiddleCenter, ContentAlignment.MiddleRight
                Y = Container.Top + (Container.Height - Size.Height) \ 2
            Case ContentAlignment.BottomLeft, ContentAlignment.BottomCenter, ContentAlignment.BottomRight
                Y = Container.Bottom - Size.Height
            Case Else
                Y = Container.Top
        End Select
        Return New Rectangle(X, Y, Math.Max(0, Size.Width), Math.Max(0, Size.Height))
    End Function
    Private Shared Function CreateRoundedRectangle(Rectangle As RectangleF, Radius As Single) As GraphicsPath
        Dim Path As New GraphicsPath()
        If Rectangle.Width <= 0.0F OrElse Rectangle.Height <= 0.0F Then Return Path
        Radius = Math.Max(0.0F, Math.Min(Radius, Math.Min(Rectangle.Width, Rectangle.Height) / 2.0F))
        If Radius <= 0.0F Then
            Path.AddRectangle(Rectangle)
            Path.CloseFigure()
            Return Path
        End If
        Dim Diameter As Single = Radius * 2.0F
        Path.AddArc(Rectangle.Left, Rectangle.Top, Diameter, Diameter, 180.0F, 90.0F)
        Path.AddArc(Rectangle.Right - Diameter, Rectangle.Top, Diameter, Diameter, 270.0F, 90.0F)
        Path.AddArc(Rectangle.Right - Diameter, Rectangle.Bottom - Diameter, Diameter, Diameter, 0.0F, 90.0F)
        Path.AddArc(Rectangle.Left, Rectangle.Bottom - Diameter, Diameter, Diameter, 90.0F, 90.0F)
        Path.CloseFigure()
        Return Path
    End Function
    Private Shared Function BlendColor(BaseColor As Color, OverlayColor As Color, Amount As Single) As Color
        Amount = Clamp(Amount, 0.0F, 1.0F)
        Dim Alpha As Integer = CInt(Math.Round(BaseColor.A + (OverlayColor.A - BaseColor.A) * Amount))
        Dim Red As Integer = CInt(Math.Round(BaseColor.R + (OverlayColor.R - BaseColor.R) * Amount))
        Dim Green As Integer = CInt(Math.Round(BaseColor.G + (OverlayColor.G - BaseColor.G) * Amount))
        Dim Blue As Integer = CInt(Math.Round(BaseColor.B + (OverlayColor.B - BaseColor.B) * Amount))
        Return Color.FromArgb(Clamp(Alpha, 0, 255), Clamp(Red, 0, 255), Clamp(Green, 0, 255), Clamp(Blue, 0, 255))
    End Function
    Private Shared Function ApplyEasing(Value As Single, Easing As ToggleButtonAnimationEasing) As Single
        Value = Clamp(Value, 0.0F, 1.0F)
        Select Case Easing
            Case ToggleButtonAnimationEasing.Linear
                Return Value
            Case ToggleButtonAnimationEasing.EaseInOutCubic
                If Value < 0.5F Then Return 4.0F * Value * Value * Value
                Return 1.0F - CSng(Math.Pow(-2.0F * Value + 2.0F, 3.0F)) / 2.0F
            Case Else
                Return 1.0F - CSng(Math.Pow(1.0F - Value, 3.0F))
        End Select
    End Function
    Private Shared Function Clamp(Value As Single, Minimum As Single, Maximum As Single) As Single
        Return Math.Max(Minimum, Math.Min(Maximum, Value))
    End Function
    Private Shared Function Clamp(Value As Integer, Minimum As Integer, Maximum As Integer) As Integer
        Return Math.Max(Minimum, Math.Min(Maximum, Value))
    End Function
    Private Sub ResetCheckedStyle()
        _CheckedStyle.SetValues(Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), Color.White, Color.FromArgb(219, 234, 254), Color.FromArgb(31, 41, 55), Color.White, Color.FromArgb(37, 99, 235))
    End Sub
    Private Function ShouldSerializeCheckedStyle() As Boolean
        Return Not _CheckedStyle.Matches(Color.FromArgb(37, 99, 235), Color.FromArgb(29, 78, 216), Color.White, Color.FromArgb(219, 234, 254), Color.FromArgb(31, 41, 55), Color.White, Color.FromArgb(37, 99, 235))
    End Function
    Private Sub ResetUncheckedStyle()
        _UncheckedStyle.SetValues(Color.FromArgb(208, 213, 221), Color.FromArgb(184, 192, 204), Color.White, Color.FromArgb(226, 229, 234), Color.FromArgb(52, 64, 84), Color.FromArgb(102, 112, 133), Color.FromArgb(102, 112, 133))
    End Sub
    Private Function ShouldSerializeUncheckedStyle() As Boolean
        Return Not _UncheckedStyle.Matches(Color.FromArgb(208, 213, 221), Color.FromArgb(184, 192, 204), Color.White, Color.FromArgb(226, 229, 234), Color.FromArgb(52, 64, 84), Color.FromArgb(102, 112, 133), Color.FromArgb(102, 112, 133))
    End Function
    Private Sub ResetDisabledStyle()
        _DisabledStyle.SetValues(Color.FromArgb(234, 236, 240), Color.FromArgb(208, 213, 221), Color.FromArgb(249, 250, 251), Color.FromArgb(234, 236, 240), Color.FromArgb(152, 162, 179), Color.FromArgb(152, 162, 179), Color.FromArgb(152, 162, 179))
    End Sub
    Private Function ShouldSerializeDisabledStyle() As Boolean
        Return Not _DisabledStyle.Matches(Color.FromArgb(234, 236, 240), Color.FromArgb(208, 213, 221), Color.FromArgb(249, 250, 251), Color.FromArgb(234, 236, 240), Color.FromArgb(152, 162, 179), Color.FromArgb(152, 162, 179), Color.FromArgb(152, 162, 179))
    End Function
End Class