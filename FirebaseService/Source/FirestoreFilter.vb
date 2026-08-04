''' <summary>
''' Represents one field filter in a structured Cloud Firestore query.
''' </summary>
Public NotInheritable Class FirestoreFilter
    ''' <summary>
    ''' Gets the Firestore field path to compare.
    ''' </summary>
    Public ReadOnly Property Field As String
    ''' <summary>
    ''' Gets the comparison operator.
    ''' </summary>
    Public ReadOnly Property [Operator] As FirestoreOperator
    ''' <summary>
    ''' Gets the comparison value.
    ''' </summary>
    Public ReadOnly Property Value As Object
    ''' <summary>
    ''' Initializes a Firestore field filter.
    ''' </summary>
    ''' <param name="Field">The Firestore field path.</param>
    ''' <param name="Operator">The comparison operator.</param>
    ''' <param name="Value">The comparison value.</param>
    Public Sub New(Field As String, [Operator] As FirestoreOperator, Value As Object)
        If String.IsNullOrWhiteSpace(Field) Then Throw New ArgumentException("The field path cannot be empty.", NameOf(Field))
        If Not [Enum].IsDefined(GetType(FirestoreOperator), [Operator]) Then Throw New ArgumentOutOfRangeException(NameOf([Operator]))
        Me.Field = Field
        Me.Operator = [Operator]
        Me.Value = Value
    End Sub
End Class
