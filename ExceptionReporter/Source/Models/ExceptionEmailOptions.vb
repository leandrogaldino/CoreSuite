''' <summary>
''' Represents the SMTP configuration used to send exception reports by email.
''' </summary>
Public Class ExceptionEmailOptions
    ''' <summary>
    ''' Gets or sets the display name of the sender.
    ''' </summary>
    Public Property FromName As String
    ''' <summary>
    ''' Gets or sets the sender email address.
    ''' </summary>
    Public Property FromEmail As String
    ''' <summary>
    ''' Gets or sets the display name of the recipient.
    ''' </summary>
    Public Property ToName As String
    ''' <summary>
    ''' Gets or sets the recipient email address.
    ''' </summary>
    Public Property ToEmail As String
    ''' <summary>
    ''' Gets or sets the SMTP server host name.
    ''' </summary>
    Public Property Host As String
    ''' <summary>
    ''' Gets or sets the SMTP server port.
    ''' </summary>
    Public Property Port As Integer = 587
    ''' <summary>
    ''' Gets or sets the user name used for SMTP authentication.
    ''' </summary>
    ''' <remarks>
    ''' When this property is empty, <see cref="FromEmail"/> is used as the user name.
    ''' </remarks>
    Public Property UserName As String
    ''' <summary>
    ''' Gets or sets the password used for SMTP authentication.
    ''' </summary>
    Public Property Password As String
    ''' <summary>
    ''' Gets or sets whether the SMTP client authenticates before sending the report.
    ''' </summary>
    Public Property UseAuthentication As Boolean = True
    ''' <summary>
    ''' Gets or sets the SMTP connection security mode.
    ''' </summary>
    Public Property SecureSocket As ExceptionReporterSecureSocket = ExceptionReporterSecureSocket.Auto
    ''' <summary>
    ''' Gets or sets the email subject.
    ''' </summary>
    Public Property Subject As String = "Aviso de Exceção na Aplicação"
    ''' <summary>
    ''' Gets or sets whether internet availability is checked before connecting to the SMTP server.
    ''' </summary>
    Public Property CheckConnectivity As Boolean = True
End Class
