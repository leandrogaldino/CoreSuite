Imports System
Imports System.ComponentModel
Imports System.Drawing

''' <summary>Editable colors with automatic change notification.</summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public NotInheritable Class ToolStripPalette
    Public Event Changed As EventHandler

    Private _Background As Color = Color.FromArgb(248, 249, 251)
    Private _DropDownBackground As Color = Color.White
    Private _Foreground As Color = Color.FromArgb(37, 44, 56)
    Private _DisabledForeground As Color = Color.FromArgb(145, 150, 160)
    Private _Accent As Color = Color.FromArgb(35, 105, 210)
    Private _HoverBackground As Color = Color.FromArgb(228, 236, 248)
    Private _PressedBackground As Color = Color.FromArgb(205, 222, 247)
    Private _CheckedBackground As Color = Color.FromArgb(207, 224, 251)
    Private _CheckedForeground As Color = Color.White
    Private _Border As Color = Color.FromArgb(213, 219, 228)
    Private _DropDownBorder As Color = Color.FromArgb(193, 202, 216)
    Private _Separator As Color = Color.FromArgb(215, 222, 231)
    Private _ImageMarginBackground As Color = Color.FromArgb(247, 249, 252)
    Private _Grip As Color = Color.FromArgb(130, 143, 162)

    Private Sub SetColor(ByRef Field As Color, Value As Color)
        If Field.ToArgb() = Value.ToArgb() Then Return
        Field = Value
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub

    <Category("Colors")>
    Public Property Background As Color
        Get
            Return _Background
        End Get
        Set(value As Color)
            SetColor(_Background, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property DropDownBackground As Color
        Get
            Return _DropDownBackground
        End Get
        Set(value As Color)
            SetColor(_DropDownBackground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property Foreground As Color
        Get
            Return _Foreground
        End Get
        Set(value As Color)
            SetColor(_Foreground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property DisabledForeground As Color
        Get
            Return _disabledForeground
        End Get
        Set(value As Color)
            SetColor(_disabledForeground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property Accent As Color
        Get
            Return _accent
        End Get
        Set(value As Color)
            SetColor(_accent, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property HoverBackground As Color
        Get
            Return _hoverBackground
        End Get
        Set(value As Color)
            SetColor(_hoverBackground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property PressedBackground As Color
        Get
            Return _pressedBackground
        End Get
        Set(value As Color)
            SetColor(_pressedBackground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property CheckedBackground As Color
        Get
            Return _checkedBackground
        End Get
        Set(value As Color)
            SetColor(_checkedBackground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property CheckedForeground As Color
        Get
            Return _checkedForeground
        End Get
        Set(value As Color)
            SetColor(_checkedForeground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property Border As Color
        Get
            Return _border
        End Get
        Set(value As Color)
            SetColor(_border, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property DropDownBorder As Color
        Get
            Return _dropDownBorder
        End Get
        Set(value As Color)
            SetColor(_dropDownBorder, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property Separator As Color
        Get
            Return _separator
        End Get
        Set(value As Color)
            SetColor(_separator, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property ImageMarginBackground As Color
        Get
            Return _imageMarginBackground
        End Get
        Set(value As Color)
            SetColor(_imageMarginBackground, value)
        End Set
    End Property

    <Category("Colors")>
    Public Property Grip As Color
        Get
            Return _grip
        End Get
        Set(value As Color)
            SetColor(_grip, value)
        End Set
    End Property

    Public Function Copy() As ToolStripPalette
        Dim Result As New ToolStripPalette With {
            ._Background = _Background,
            ._DropDownBackground = _DropDownBackground,
            ._Foreground = _Foreground,
            ._DisabledForeground = _DisabledForeground,
            ._Accent = _Accent,
            ._HoverBackground = _HoverBackground,
            ._PressedBackground = _PressedBackground,
            ._CheckedBackground = _CheckedBackground,
            ._CheckedForeground = _CheckedForeground,
            ._Border = _Border,
            ._DropDownBorder = _DropDownBorder,
            ._Separator = _Separator,
            ._ImageMarginBackground = _ImageMarginBackground,
            ._Grip = _Grip
        }
        Return Result
    End Function

    Public Shared Function Create(Theme As ToolStripTheme) As ToolStripPalette
        Dim Palette As New ToolStripPalette()
        Select Case Theme
            Case ToolStripTheme.Dark
                Palette._Background = Color.FromArgb(37, 40, 47)
                Palette._DropDownBackground = Color.FromArgb(44, 47, 55)
                Palette._Foreground = Color.FromArgb(239, 241, 245)
                Palette._DisabledForeground = Color.FromArgb(125, 131, 142)
                Palette._Accent = Color.FromArgb(107, 167, 249)
                Palette._HoverBackground = Color.FromArgb(67, 74, 88)
                Palette._PressedBackground = Color.FromArgb(85, 96, 113)
                Palette._CheckedBackground = Color.FromArgb(65, 90, 132)
                Palette._CheckedForeground = Color.FromArgb(255, 255, 255)
                Palette._Border = Color.FromArgb(68, 74, 84)
                Palette._DropDownBorder = Color.FromArgb(91, 100, 115)
                Palette._Separator = Color.FromArgb(83, 90, 103)
                Palette._ImageMarginBackground = Color.FromArgb(39, 42, 50)
                Palette._Grip = Color.FromArgb(148, 153, 164)
            Case ToolStripTheme.Minimal
                Palette._Background = Color.FromArgb(255, 255, 255)
                Palette._DropDownBackground = Color.FromArgb(255, 255, 255)
                Palette._Foreground = Color.FromArgb(33, 36, 42)
                Palette._DisabledForeground = Color.FromArgb(170, 170, 170)
                Palette._Accent = Color.FromArgb(52, 52, 52)
                Palette._HoverBackground = Color.FromArgb(241, 242, 244)
                Palette._PressedBackground = Color.FromArgb(227, 229, 232)
                Palette._CheckedBackground = Color.FromArgb(234, 235, 237)
                Palette._CheckedForeground = Color.FromArgb(255, 255, 255)
                Palette._Border = Color.FromArgb(228, 229, 232)
                Palette._DropDownBorder = Color.FromArgb(220, 222, 226)
                Palette._Separator = Color.FromArgb(233, 234, 236)
                Palette._ImageMarginBackground = Color.FromArgb(255, 255, 255)
                Palette._Grip = Color.FromArgb(175, 177, 180)
            Case ToolStripTheme.System
                Palette._Background = SystemColors.MenuBar
                Palette._DropDownBackground = SystemColors.Menu
                Palette._Foreground = SystemColors.MenuText
                Palette._DisabledForeground = SystemColors.GrayText
                Palette._Accent = SystemColors.Highlight
                Palette._HoverBackground = SystemColors.Highlight
                Palette._PressedBackground = SystemColors.Highlight
                Palette._CheckedBackground = SystemColors.Highlight
                Palette._CheckedForeground = SystemColors.HighlightText
                Palette._Border = SystemColors.ControlDark
                Palette._DropDownBorder = SystemColors.ControlDark
                Palette._Separator = SystemColors.ControlDark
                Palette._ImageMarginBackground = SystemColors.Menu
                Palette._Grip = SystemColors.ControlDark
        End Select
        Return Palette
    End Function

    Public Overrides Function ToString() As String
        Return "ToolStripPalette (customizable)"
    End Function
End Class
