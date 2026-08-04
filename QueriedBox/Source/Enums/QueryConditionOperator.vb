''' <summary>
''' Specifies the comparison operator used in query conditions.
''' </summary>
Public Enum QueryConditionOperator
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
    ''' <summary>
    ''' Compares whether a value is within a specified range using the SQL BETWEEN operator.
    ''' </summary>
    Between
    ''' <summary>
    ''' Compares whether a value exists within a specified collection using the SQL IN operator.
    ''' </summary>
    [In]
    ''' <summary>
    ''' Compares whether a value does not exist within a specified collection using the SQL NOT IN operator.
    ''' </summary>
    NotIn
End Enum