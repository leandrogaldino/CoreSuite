''' <summary>
''' Describes the result of a password hash verification operation.
''' </summary>
Public Enum PasswordVerificationResult
    ''' <summary>
    ''' The password did not match or the stored hash was malformed.
    ''' </summary>
    Failed = 0
    ''' <summary>
    ''' The password matched and the stored hash uses the current configuration.
    ''' </summary>
    Success = 1
    ''' <summary>
    ''' The password matched, but the stored hash should be replaced with a newly generated hash.
    ''' </summary>
    SuccessRehashNeeded = 2
    End Enum