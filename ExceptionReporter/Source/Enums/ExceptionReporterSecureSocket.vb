''' <summary>
''' Defines how an SMTP connection is secured when an exception report is sent.
''' </summary>
Public Enum ExceptionReporterSecureSocket
    ''' <summary>
    ''' Selects the security mode automatically from the SMTP port and server capabilities.
    ''' </summary>
    Auto
    ''' <summary>
    ''' Requires the connection to be upgraded to TLS using STARTTLS.
    ''' </summary>
    StartTls
    ''' <summary>
    ''' Uses STARTTLS when the server supports it and otherwise continues without TLS.
    ''' </summary>
    StartTlsWhenAvailable
    ''' <summary>
    ''' Establishes a TLS connection immediately.
    ''' </summary>
    SslOnConnect
    ''' <summary>
    ''' Establishes an unencrypted SMTP connection.
    ''' </summary>
    None
End Enum
