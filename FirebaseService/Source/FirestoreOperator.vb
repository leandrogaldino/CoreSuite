''' <summary>
''' Defines supported Firestore query operators.
''' </summary>
Public Enum FirestoreOperator

    ''' <summary>Equal comparison.</summary>
    Equal

    ''' <summary>Not equal comparison.</summary>
    NotEqual

    ''' <summary>Less than comparison.</summary>
    LessThan

    ''' <summary>Less than or equal comparison.</summary>
    LessThanOrEqual

    ''' <summary>Greater than comparison.</summary>
    GreaterThan

    ''' <summary>Greater than or equal comparison.</summary>
    GreaterThanOrEqual

    ''' <summary>Checks if an array contains a value.</summary>
    ArrayContains

    ''' <summary>Checks if an array contains any value from a list.</summary>
    ArrayContainsAny

    ''' <summary>Checks if a value exists within a list.</summary>
    InList

    ''' <summary>Checks if a value does not exist within a list.</summary>
    NotInList

End Enum