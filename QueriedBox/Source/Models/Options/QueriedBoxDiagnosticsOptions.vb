Imports System.ComponentModel
''' <summary>
''' Provides diagnostic options for query execution performed by a <see cref="QueriedBox"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxDiagnosticsOptions
    ''' <summary>
    ''' Gets or sets whether the configured query is debugged when the text changes.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets whether the configured query is debugged when the text changes.")>
    Public Property DebugOnTextChanged As Boolean
    ''' <summary>
    ''' Returns a summary of the configured diagnostic options.
    ''' </summary>
    ''' <returns>
    ''' A string indicating whether query debugging is enabled.
    ''' </returns>
    Public Overrides Function ToString() As String
        Return If(DebugOnTextChanged, "Debug enabled", "Debug disabled")
    End Function
End Class

