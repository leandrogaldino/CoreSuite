Imports System.ComponentModel
Imports System.Drawing

''' <summary>Defines the colors for one visual state of the toggle.</summary>
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
