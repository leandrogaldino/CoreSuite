''' <summary>
''' Defines the security mode used when establishing a connection with the mail server.
''' </summary>
''' <remarks>
''' Specifies how the SMTP connection should handle SSL/TLS encryption during the
''' connection process.
''' </remarks>
Public Enum CMessageBoxSecureSocket

    ''' <summary>
    ''' Automatically determines the appropriate security mode based on the server capabilities.
    ''' </summary>
    Auto

    ''' <summary>
    ''' Establishes an unencrypted connection and then upgrades it to a secure connection
    ''' using the STARTTLS command.
    ''' </summary>
    StartTls

    ''' <summary>
    ''' Uses STARTTLS encryption when supported by the server; otherwise, continues
    ''' with an unencrypted connection.
    ''' </summary>
    StartTlsWhenAvailable

    ''' <summary>
    ''' Establishes a secure SSL/TLS connection immediately when connecting to the server.
    ''' </summary>
    SslOnConnect

    ''' <summary>
    ''' Disables SSL/TLS encryption and establishes an unsecured connection.
    ''' </summary>
    None

End Enum