''' <summary>
''' Specifies the logical relation used to combine query conditions.
''' </summary>
Public Enum QueryRelation
    ''' <summary>
    ''' Combines conditions using the logical AND operator.
    ''' All combined conditions must be satisfied.
    ''' </summary>
    [And]
    ''' <summary>
    ''' Combines conditions using the logical OR operator.
    ''' At least one combined condition must be satisfied.
    ''' </summary>
    [Or]
End Enum