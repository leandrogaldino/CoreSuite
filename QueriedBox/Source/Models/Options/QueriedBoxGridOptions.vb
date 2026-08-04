Imports System.ComponentModel
''' <summary>
''' Provides appearance settings for the results grid displayed by a <see cref="QueriedBox"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxGridOptions
    Private _BackColor As Color = SystemColors.Window
    Private _ForeColor As Color = SystemColors.ControlText
    Private _SelectionBackColor As Color = SystemColors.HotTrack
    Private _SelectionForeColor As Color = SystemColors.Window
    Private _HeaderBackColor As Color = SystemColors.Window
    Private _HeaderForeColor As Color = SystemColors.ControlText
    ''' <summary>
    ''' Gets or sets the background color of the results grid.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "Window")>
    <Description("Gets or sets the background color of the results grid.")>
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _BackColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text color of the results grid.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Gets or sets the text color of the results grid.")>
    Public Property ForeColor As Color
        Get
            Return _ForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _ForeColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selection background color of the results grid.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "HotTrack")>
    <Description("Gets or sets the selection background color of the results grid.")>
    Public Property SelectionBackColor As Color
        Get
            Return _SelectionBackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _SelectionBackColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the selection text color of the results grid.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "Window")>
    <Description("Gets or sets the selection text color of the results grid.")>
    Public Property SelectionForeColor As Color
        Get
            Return _SelectionForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _SelectionForeColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property
    ''' <summary> 
    ''' Gets Or sets the background color of the grid headers.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "Control")>
    <Description("Gets or sets the background color of the grid headers.")>
    Public Property HeaderBackColor As Color
        Get
            Return _HeaderBackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _HeaderBackColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text color of the grid headers.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Gets or sets the text color of the grid headers.")>
    Public Property HeaderForeColor As Color
        Get
            Return _HeaderForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _HeaderForeColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether the grid headers use bold font.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Gets or sets whether the grid headers use bold font.")>
    Public Property HeadersBold As Boolean = False
    ''' <summary>
    ''' Gets or sets whether the grid headers are visible.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Gets or sets whether the grid headers are visible.")>
    Public Property HeaderVisible As Boolean = True
    ''' <summary>
    ''' Gets or sets whether the results grid displays vertical grid lines.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Gets or sets whether the results grid displays vertical grid lines.")>
    Public Property ShowVerticalLines As Boolean = True

    ''' <summary>
    ''' Returns a summary of the configured results grid appearance options.
    ''' </summary>
    ''' <returns>
    ''' A string describing the main colors, header visibility, and vertical
    ''' grid line configuration.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim headerState As String = If(HeaderVisible, "Visible", "Hidden")
        Dim headerStyle As String = If(HeadersBold, "Bold", "Regular")
        Dim verticalLinesState As String = If(ShowVerticalLines, "Visible", "Hidden")
        Return $"Background: {BackColor.Name}, " & $"Selection: {SelectionBackColor.Name}, " & $"Headers: {headerState}/{headerStyle}, " & $"Vertical lines: {verticalLinesState}"
    End Function
End Class
