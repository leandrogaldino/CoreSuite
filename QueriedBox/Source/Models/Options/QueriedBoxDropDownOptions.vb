Imports System.ComponentModel
''' <summary>
''' Provides configuration options for the results dropdown displayed by a <see cref="QueriedBox"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxDropDownOptions
    Private _BorderColor As Color = SystemColors.HotTrack
    Private _AutoStretchRight As Boolean
    Private _StretchRight As Integer
    ''' <summary>
    ''' Gets or sets the border color of the results dropdown.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "HotTrack")>
    <Description("Gets or sets the border color of the results dropdown.")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _BorderColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the dropdown automatically expands horizontally to display all results.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Gets or sets whether the dropdown automatically expands horizontally to display all results.")>
    Public Property AutoStretchRight As Boolean
        Get
            Return _AutoStretchRight
        End Get
        Set(value As Boolean)
            _AutoStretchRight = value
            If value Then StretchRight = 0
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the amount of vertical expansion applied to the dropdown.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Integer), "120")>
    <Description("Gets or sets the amount of vertical expansion applied to the dropdown.")>
    Public Property StretchDown As Integer = 120
    ''' <summary>
    ''' Gets or sets the amount of horizontal expansion applied to the left side of the dropdown.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Integer), "0")>
    <Description("Gets or sets the amount of horizontal expansion applied to the left side of the dropdown.")>
    Public Property StretchLeft As Integer
    ''' <summary>
    ''' Gets or sets the amount of horizontal expansion applied to the right side of the dropdown.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Integer), "0")>
    <Description("Gets or sets the amount of horizontal expansion applied to the right side of the dropdown.")>
    Public Property StretchRight As Integer
        Get
            Return _StretchRight
        End Get
        Set(value As Integer)
            If AutoStretchRight Then value = 0
            _StretchRight = value
        End Set
    End Property
    ''' <summary>
    ''' Returns a summary of the configured dropdown layout options.
    ''' </summary>
    ''' <returns>
    ''' A string describing the dropdown stretch configuration.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim rightValue As String =
            If(AutoStretchRight, "Auto", StretchRight.ToString())

        Return $"Left: {StretchLeft}, Right: {rightValue}, Down: {StretchDown}"
    End Function
End Class
