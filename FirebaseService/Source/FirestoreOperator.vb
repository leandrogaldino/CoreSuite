''' <summary>
''' Specifies the comparison operators available for building Firestore query filters.
''' </summary>
Public Enum FirestoreOperator
    ''' <summary>
    ''' Matches values that are equal to the specified value.
    ''' </summary>
    Equal
    ''' <summary>
    ''' Matches values that are not equal to the specified value.
    ''' </summary>
    NotEqual
    ''' <summary>
    ''' Matches values that are less than the specified value.
    ''' </summary>
    LessThan
    ''' <summary>
    ''' Matches values that are less than or equal to the specified value.
    ''' </summary>
    LessThanOrEqual
    ''' <summary>
    ''' Matches values that are greater than the specified value.
    ''' </summary>
    GreaterThan
    ''' <summary>
    ''' Matches values that are greater than or equal to the specified value.
    ''' </summary>
    GreaterThanOrEqual
    ''' <summary>
    ''' Matches array fields that contain the specified value.
    ''' </summary>
    ArrayContains
    ''' <summary>
    ''' Matches array fields that contain any of the specified values.
    ''' </summary>
    ArrayContainsAny
    ''' <summary>
    ''' Matches values that are present in the specified list of values.
    ''' </summary>
    InList
    ''' <summary>
    ''' Matches values that are not present in the specified list of values.
    ''' </summary>
    NotInList
End Enum