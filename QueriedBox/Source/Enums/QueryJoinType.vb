''' <summary>
''' Specifies the type of JOIN operation used when combining tables in a query.
''' </summary>
Public Enum QueryJoinType
    ''' <summary>
    ''' Combines records when matching values exist in both tables.
    ''' </summary>
    Inner
    ''' <summary>
    ''' Returns all records from the left table and matching records from the right table.
    ''' </summary>
    Left
    ''' <summary>
    ''' Returns all records from the right table and matching records from the left table.
    ''' </summary>
    Right
    ''' <summary>
    ''' Returns all records from both tables, including unmatched records.
    ''' </summary>
    Full
End Enum