''' <summary>
''' Specifies how entered text is matched against values during local filtering.
''' </summary>
Public Enum DataGridViewFilterSearchMode
    ''' <summary>
    ''' Matches values that contain the entered text anywhere.
    ''' </summary>
    Contains
    ''' <summary>
    ''' Matches values that begin with the entered text.
    ''' </summary>
    StartsWith
    ''' <summary>
    ''' Matches values that end with the entered text.
    ''' </summary>
    EndsWith
    ''' <summary>
    ''' Matches values whose complete textual representation equals the entered text.
    ''' </summary>
    ExactMatch
End Enum
