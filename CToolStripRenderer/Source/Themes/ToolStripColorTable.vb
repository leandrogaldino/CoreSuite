Imports System.Drawing
Imports System.Windows.Forms

''' <summary>Fallback colors for framework-rendered ToolStrip parts.</summary>
Friend NotInheritable Class ToolStripColorTable
    Inherits ProfessionalColorTable

    Private _Palette As ToolStripPalette = ToolStripPalette.Create(ToolStripTheme.Light)

    Friend Sub SetPalette(palette As ToolStripPalette)
        _Palette = palette
        Me.UseSystemColors = False
    End Sub

    Public Overrides ReadOnly Property ToolStripGradientBegin As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property ToolStripGradientMiddle As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property ToolStripGradientEnd As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property MenuStripGradientBegin As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property MenuStripGradientEnd As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property StatusStripGradientBegin As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property StatusStripGradientEnd As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
        Get
            Return _Palette.DropDownBackground
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelected As Color
        Get
            Return _Palette.HoverBackground
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelectedGradientBegin As Color
        Get
            Return _Palette.HoverBackground
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelectedGradientEnd As Color
        Get
            Return _Palette.HoverBackground
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemBorder As Color
        Get
            Return _Palette.Border
        End Get
    End Property

    Public Overrides ReadOnly Property MenuBorder As Color
        Get
            Return _Palette.DropDownBorder
        End Get
    End Property

    Public Overrides ReadOnly Property ToolStripBorder As Color
        Get
            Return _Palette.Border
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonSelectedGradientBegin As Color
        Get
            Return _Palette.HoverBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonSelectedGradientMiddle As Color
        Get
            Return _Palette.HoverBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonSelectedGradientEnd As Color
        Get
            Return _Palette.HoverBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonSelectedBorder As Color
        Get
            Return _Palette.Border
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonPressedGradientBegin As Color
        Get
            Return _Palette.PressedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonPressedGradientMiddle As Color
        Get
            Return _Palette.PressedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonPressedGradientEnd As Color
        Get
            Return _Palette.PressedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonPressedBorder As Color
        Get
            Return _Palette.Border
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonCheckedGradientBegin As Color
        Get
            Return _Palette.CheckedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonCheckedGradientMiddle As Color
        Get
            Return _Palette.CheckedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ButtonCheckedGradientEnd As Color
        Get
            Return _Palette.CheckedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property CheckBackground As Color
        Get
            Return _Palette.CheckedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property CheckSelectedBackground As Color
        Get
            Return _Palette.CheckedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property CheckPressedBackground As Color
        Get
            Return _Palette.CheckedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property SeparatorDark As Color
        Get
            Return _Palette.Separator
        End Get
    End Property

    Public Overrides ReadOnly Property SeparatorLight As Color
        Get
            Return _Palette.Separator
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
        Get
            Return _Palette.ImageMarginBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
        Get
            Return _Palette.ImageMarginBackground
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
        Get
            Return _Palette.ImageMarginBackground
        End Get
    End Property

    Public Overrides ReadOnly Property GripDark As Color
        Get
            Return _Palette.Grip
        End Get
    End Property

    Public Overrides ReadOnly Property GripLight As Color
        Get
            Return _Palette.Background
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemPressedGradientBegin As Color
        Get
            Return _Palette.PressedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemPressedGradientMiddle As Color
        Get
            Return _Palette.PressedBackground
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemPressedGradientEnd As Color
        Get
            Return _Palette.PressedBackground
        End Get
    End Property

End Class
