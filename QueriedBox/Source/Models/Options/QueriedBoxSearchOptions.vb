Imports System.ComponentModel
''' <summary>
''' Provides search behavior settings for a <see cref="QueriedBox"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxSearchOptions
    Private _MinimumChars As Integer = 3
    Private _Interval As Integer = 300
    Private _Enabled As Boolean = True
    ''' <summary>
    ''' Gets or sets the minimum number of characters required to start a query.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Integer), "3")>
    <Description("Gets or sets the minimum number of characters required to start a query.")>
    Public Property MinimumChars As Integer
        Get
            Return _MinimumChars
        End Get
        Set(value As Integer)
            If value < 1 Then value = 1
            _MinimumChars = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether query functionality is enabled.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Gets or sets whether query functionality is enabled.")>
    Public Property Enabled As Boolean
        Get
            Return _Enabled
        End Get
        Set(value As Boolean)
            _Enabled = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the delay interval between queries while typing.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Integer), "300")>
    <Description("Gets or sets the delay interval between queries while typing.")>
    Public Property Interval As Integer
        Get
            Return _Interval
        End Get
        Set(value As Integer)
            _Interval = value
        End Set
    End Property
    ''' <summary>
    ''' Returns a summary of the configured search behavior options.
    ''' </summary>
    ''' <returns>
    ''' A string describing whether searching is enabled, the minimum number
    ''' of required characters, and the query delay interval.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim state As String =
            If(Enabled, "Enabled", "Disabled")

        Return $"{state}, Minimum characters: {MinimumChars}, Interval: {Interval} ms"
    End Function
End Class
