Imports System.Collections.ObjectModel
''' <summary>
''' Provides data for the <see cref="ValidationProvider.ValidationCompleted"/> event.
''' </summary>
Public NotInheritable Class ValidationCompletedEventArgs
    Inherits EventArgs
    Private ReadOnly _Results As ReadOnlyCollection(Of ValidationResult)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidationCompletedEventArgs"/> class.
    ''' </summary>
    ''' <param name="Results">The results produced by the validation operation.</param>
    Public Sub New(Results As IEnumerable(Of ValidationResult))
        If Results Is Nothing Then Throw New ArgumentNullException(NameOf(Results))
        _Results = New List(Of ValidationResult)(Results).AsReadOnly()
    End Sub
    ''' <summary>
    ''' Gets all results produced by the validation operation.
    ''' </summary>
    ''' <value>A read-only list containing one result for each evaluated control.</value>
    Public ReadOnly Property Results As IReadOnlyList(Of ValidationResult)
        Get
            Return _Results
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether every evaluated control is valid.
    ''' </summary>
    ''' <value><see langword="True"/> when no result failed; otherwise, <see langword="False"/>.</value>
    Public ReadOnly Property IsValid As Boolean
        Get
            For Each Result As ValidationResult In _Results
                If Not Result.IsValid Then Return False
            Next
            Return True
        End Get
    End Property
    ''' <summary>
    ''' Gets the number of controls that failed validation.
    ''' </summary>
    ''' <value>The count of invalid validation results.</value>
    Public ReadOnly Property InvalidControlCount As Integer
        Get
            Dim Count As Integer
            For Each Result As ValidationResult In _Results
                If Not Result.IsValid Then Count += 1
            Next
            Return Count
        End Get
    End Property
End Class
