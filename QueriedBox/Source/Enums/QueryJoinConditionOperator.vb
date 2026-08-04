''' <summary>
''' Specifies the comparison operator used in query join conditions.
''' </summary>
Public Enum QueryJoinConditionOperator
    ''' <summary>
    ''' Compares values for equality.
    ''' </summary>
    Equal
    ''' <summary>
    ''' Compares values for inequality.
    ''' </summary>
    NotEqual
    ''' <summary>
    ''' Compares whether a value is greater than another value.
    ''' </summary>
    GreaterThan
    ''' <summary>
    ''' Compares whether a value is less than another value.
    ''' </summary>
    LessThan
    ''' <summary>
    ''' Compares whether a value is greater than or equal to another value.
    ''' </summary>
    GreaterThanOrEqual
    ''' <summary>
    ''' Compares whether a value is less than or equal to another value.
    ''' </summary>
    LessThanOrEqual
    ''' <summary>
    ''' Performs a pattern matching comparison using the SQL LIKE operator.
    ''' </summary>
    [Like]
End Enum