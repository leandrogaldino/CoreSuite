''' <summary>
''' Provides data for the <see cref="DataGridViewFilterBox.FilterFailed"/> event.
''' </summary>
Public Class FilterFailedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FilterFailedEventArgs"/> class.
    ''' </summary>
    ''' <param name="FilterText">The filter text being processed when the failure occurred.</param>
    ''' <param name="Failure">The exception that describes the failure.</param>
    Public Sub New(FilterText As String, Failure As Exception)
        Me.FilterText = FilterText
        Me.Exception = Failure
    End Sub
    ''' <summary>
    ''' Gets the filter text being processed when the failure occurred.
    ''' </summary>
    ''' <value>The current text entered in the filter box.</value>
    Public ReadOnly Property FilterText As String
    ''' <summary>
    ''' Gets the exception that describes why the filter could not be applied.
    ''' </summary>
    ''' <value>The exception raised while resolving or applying the filter.</value>
    Public ReadOnly Property Exception As Exception
End Class
