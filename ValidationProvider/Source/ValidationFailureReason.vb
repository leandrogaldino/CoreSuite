''' <summary>
''' Identifies the rule that caused a control to fail validation.
''' </summary>
Public Enum ValidationFailureReason
    ''' <summary>
    ''' Indicates that the control is valid or that no built-in rule failed.
    ''' </summary>
    None
    ''' <summary>
    ''' Indicates that a required control does not contain a value.
    ''' </summary>
    Required
    ''' <summary>
    ''' Indicates that the represented text is shorter than the configured minimum length.
    ''' </summary>
    MinimumLength
    ''' <summary>
    ''' Indicates that the represented text is longer than the configured maximum length.
    ''' </summary>
    MaximumLength
    ''' <summary>
    ''' Indicates that the represented text does not match the configured regular expression or complete mask.
    ''' </summary>
    InvalidFormat
    ''' <summary>
    ''' Indicates that the represented value does not match the comparison control.
    ''' </summary>
    Comparison
    ''' <summary>
    ''' Indicates that an application-defined validation rule rejected the value.
    ''' </summary>
    Custom
End Enum
