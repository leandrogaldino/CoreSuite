''' <summary>
''' Represents a Firestore query filter.
''' </summary>
Public Class FirestoreFilter

    ''' <summary>
    ''' Gets or sets the field name to be filtered.
    ''' </summary>
    Public Property Field As String

    ''' <summary>
    ''' Gets or sets the comparison operator.
    ''' </summary>
    Public Property [Operator] As FirestoreOperator

    ''' <summary>
    ''' Gets or sets the comparison value.
    ''' </summary>
    Public Property Value As Object

    ''' <summary>
    ''' Creates a new Firestore filter.
    ''' </summary>
    ''' <param name="Field">Field name.</param>
    ''' <param name="Operator">Comparison operator.</param>
    ''' <param name="Value">Comparison value.</param>
    Public Sub New(Field As String, [Operator] As FirestoreOperator, Value As Object)
        Me.Field = Field
        Me.Operator = [Operator]
        Me.Value = Value
    End Sub

End Class