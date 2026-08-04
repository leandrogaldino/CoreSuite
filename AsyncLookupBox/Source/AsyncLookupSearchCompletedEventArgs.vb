''' <summary>
''' Provides data for the <see cref="AsyncLookupBox.SearchCompleted"/> event.
''' </summary>
Public NotInheritable Class AsyncLookupSearchCompletedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupSearchCompletedEventArgs"/> class.
    ''' </summary>
    ''' <param name="SearchText">The text used by the completed search.</param>
    ''' <param name="Results">The results retained by the control.</param>
    ''' <param name="Duration">The elapsed search duration.</param>
    ''' <param name="WasTruncated">Indicates whether results beyond <see cref="AsyncLookupBox.MaximumResults"/> were discarded.</param>
    Public Sub New(SearchText As String, Results As IReadOnlyList(Of Object), Duration As TimeSpan, WasTruncated As Boolean)
        Me.SearchText = If(SearchText, String.Empty)
        Me.Results = If(Results, Array.Empty(Of Object)())
        Me.Duration = Duration
        Me.WasTruncated = WasTruncated
    End Sub
    ''' <summary>
    ''' Gets the text used by the completed search.
    ''' </summary>
    Public ReadOnly Property SearchText As String
    ''' <summary>
    ''' Gets the results retained and displayed by the control.
    ''' </summary>
    Public ReadOnly Property Results As IReadOnlyList(Of Object)
    ''' <summary>
    ''' Gets the number of retained results.
    ''' </summary>
    Public ReadOnly Property ResultCount As Integer
        Get
            Return Results.Count
        End Get
    End Property
    ''' <summary>
    ''' Gets the elapsed duration of the search request.
    ''' </summary>
    Public ReadOnly Property Duration As TimeSpan
    ''' <summary>
    ''' Gets a value indicating whether results beyond the configured maximum were discarded.
    ''' </summary>
    Public ReadOnly Property WasTruncated As Boolean
End Class
