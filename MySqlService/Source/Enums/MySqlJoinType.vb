''' <summary>
''' Specifies the SQL join operation used by a structured SELECT.
''' </summary>
Public Enum MySqlJoinType
    ''' <summary>
    ''' Returns rows that have matching values in both sources.
    ''' </summary>
    Inner
    ''' <summary>
    ''' Returns all rows from the left source and matching rows from the joined source.
    ''' </summary>
    Left
    ''' <summary>
    ''' Returns all rows from the joined source and matching rows from the left source.
    ''' </summary>
    Right
    ''' <summary>
    ''' Returns the Cartesian product of both sources and does not use an ON condition.
    ''' </summary>
    Cross
End Enum
