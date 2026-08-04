''' <summary>
''' Provides data for the <see cref="AsyncLookupBox.SearchFailed"/> event.
''' </summary>
Public NotInheritable Class AsyncLookupSearchFailedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupSearchFailedEventArgs"/> class.
    ''' </summary>
    ''' <param name="SearchText">The text used by the failed search.</param>
    ''' <param name="Exception">The exception raised while preparing or awaiting the search operation.</param>
    Public Sub New(SearchText As String, Exception As Exception)
        ArgumentNullException.ThrowIfNull(Exception)
        Me.SearchText = If(SearchText, String.Empty)
        Me.Exception = Exception
    End Sub
    ''' <summary>
    ''' Gets the text used by the failed search.
    ''' </summary>
    Public ReadOnly Property SearchText As String
    ''' <summary>
    ''' Gets the exception raised by the search.
    ''' </summary>
    Public ReadOnly Property Exception As Exception
End Class
