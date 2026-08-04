Imports System.Collections.ObjectModel
Imports System.Threading
''' <summary>
''' Provides data for the <see cref="DataGridViewFilterBox.FilterRequested"/> event.
''' </summary>
Public Class FilterRequestedEventArgs
    Inherits EventArgs
    Private ReadOnly _ColumnNames As ReadOnlyCollection(Of String)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FilterRequestedEventArgs"/> class.
    ''' </summary>
    ''' <param name="FilterText">The text to use in the custom or remote filter.</param>
    ''' <param name="ColumnNames">The configured column names relevant to the request.</param>
    ''' <param name="CancellationToken">A token canceled when the request is superseded or cleared.</param>
    Public Sub New(FilterText As String, ColumnNames As IEnumerable(Of String), CancellationToken As CancellationToken)
        Me.FilterText = FilterText
        _ColumnNames = New List(Of String)(ColumnNames).AsReadOnly()
        Me.CancellationToken = CancellationToken
    End Sub
    ''' <summary>
    ''' Gets the text to use in the custom or remote filter.
    ''' </summary>
    ''' <value>The current text entered in the filter box.</value>
    Public ReadOnly Property FilterText As String
    ''' <summary>
    ''' Gets the configured column names relevant to the filter request.
    ''' </summary>
    ''' <value>A read-only list containing data property or column names.</value>
    Public ReadOnly Property ColumnNames As IReadOnlyList(Of String)
        Get
            Return _ColumnNames
        End Get
    End Property
    ''' <summary>
    ''' Gets a token that is canceled when a newer request replaces this request, the filter is cleared, or the control is disposed.
    ''' </summary>
    ''' <value>The cancellation token associated with this request.</value>
    Public ReadOnly Property CancellationToken As CancellationToken
End Class
