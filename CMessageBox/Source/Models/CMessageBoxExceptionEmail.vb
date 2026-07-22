''' <summary>
''' Represents the configuration required to send exception notifications by email.
''' </summary>
''' <remarks>
''' This class contains the SMTP connection settings and sender/recipient information
''' used by the CMessageBox to send exception reports.
''' </remarks>
Public Class CMessageBoxExceptionEmail

    ''' <summary>
    ''' Gets or sets the display name of the email sender.
    ''' </summary>
    Public Property FromName As String

    ''' <summary>
    ''' Gets or sets the email address used as the sender.
    ''' </summary>
    Public Property FromEmail As String

    ''' <summary>
    ''' Gets or sets the display name of the email recipient.
    ''' </summary>
    Public Property ToName As String

    ''' <summary>
    ''' Gets or sets the email address that will receive exception notifications.
    ''' </summary>
    Public Property ToEmail As String

    ''' <summary>
    ''' Gets or sets the password used to authenticate with the SMTP server.
    ''' </summary>
    Public Property Password As String

    ''' <summary>
    ''' Gets or sets the SMTP server port used for the connection.
    ''' </summary>
    Public Property Port As Integer

    ''' <summary>
    ''' Gets or sets the SMTP server host address.
    ''' </summary>
    Public Property Host As String

    ''' <summary>
    ''' Gets or sets the security mode used when connecting to the SMTP server.
    ''' </summary>
    ''' <remarks>
    ''' The default value is <see cref="CMessageBoxSecureSocket.Auto"/>, allowing the
    ''' connection security mode to be determined automatically.
    ''' </remarks>
    Public Property SecureSocket As CMessageBoxSecureSocket = CMessageBoxSecureSocket.Auto

End Class