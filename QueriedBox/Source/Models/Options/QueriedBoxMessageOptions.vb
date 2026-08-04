Imports System.ComponentModel
''' <summary>
''' Provides message and message appearance settings for a
''' <see cref="QueriedBox"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxMessageOptions
    Private _BackColor As Color = SystemColors.Window
    Private _ForeColor As Color = SystemColors.ControlText
    ''' <summary>
    ''' Gets or sets the text displayed when no results are found.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue("No results found.")>
    <Description("Gets or sets the text displayed when no results are found.")>
    Public Property NoResultsText As String = "No results found."
    ''' <summary>
    ''' Gets or sets the text displayed when one more character is required
    ''' before starting a query.
    ''' </summary>
    ''' <remarks>
    ''' The placeholder {0} is replaced with the remaining character count.
    ''' </remarks>
    <Category("QueriedBox")>
    <DefaultValue("Type {0} more character to search.")>
    <Description("Gets or sets the text displayed when one more character is required before starting a query. The placeholder {0} is replaced with the remaining character count.")>
    Public Property CharsRemainingSingularText As String =
        "Type {0} more character to search."

    ''' <summary>
    ''' Gets or sets the text displayed when multiple additional characters
    ''' are required before starting a query.
    ''' </summary>
    ''' <remarks>
    ''' The placeholder {0} is replaced with the remaining character count.
    ''' </remarks>
    <Category("QueriedBox")>
    <DefaultValue("Type {0} more characters to search.")>
    <Description("Gets or sets the text displayed when multiple additional characters are required before starting a query. The placeholder {0} is replaced with the remaining character count.")>
    Public Property CharsRemainingPluralText As String = "Type {0} more characters to search."
    ''' <summary>
    ''' Gets or sets the background color of the label that displays message
    ''' information.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "Window")>
    <Description("Gets or sets the background color of the label that displays message information.")>
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(value As Color)
            If value = Color.Transparent Then
                Common.ThrowNoTransparentColorException()
                Return
            End If

            _BackColor = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text color of the label that displays message
    ''' information.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Gets or sets the text color of the label that displays message information.")>
    Public Property ForeColor As Color
        Get
            Return _ForeColor
        End Get
        Set(value As Color)
            If value = Color.Transparent Then
                Common.ThrowNoTransparentColorException()
                Return
            End If

            _ForeColor = value
        End Set
    End Property
    ''' <summary>
    ''' Returns a summary of the configured message options.
    ''' </summary>
    ''' <returns>
    ''' A string describing the no-results message and the configured message
    ''' colors.
    ''' </returns>
    Public Overrides Function ToString() As String
        Return $"No results: ""{NoResultsText}"", " & $"Background: {BackColor.Name}, " & $"Text: {ForeColor.Name}"
    End Function
End Class