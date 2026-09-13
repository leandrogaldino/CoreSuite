Imports System.ComponentModel
Imports System.Diagnostics

Namespace CoreSuite.OutlookEmail

    ''' <summary>
    ''' Represents an e-mail message to be composed in Microsoft Outlook Classic.
    ''' </summary>
    <DebuggerDisplay("{Subject,nq}")>
    Public Class OutlookEmailMessage
        ''' <summary>
        ''' Initializes a new instance of the <see cref="OutlookEmailMessage"/> class.
        ''' </summary>
        Public Sub New()
            ToRecipients = New List(Of String)
            CcRecipients = New List(Of String)
            BccRecipients = New List(Of String)
            Attachments = New List(Of String)
        End Sub

        ''' <summary>
        ''' Gets the recipients that will be placed in the To field.
        ''' </summary>
        <Category("Recipients")>
        <Description("Recipients placed in the To field.")>
        Public ReadOnly Property ToRecipients As List(Of String)

        ''' <summary>
        ''' Gets the recipients that will be placed in the Cc field.
        ''' </summary>
        <Category("Recipients")>
        <Description("Recipients placed in the Cc field.")>
        Public ReadOnly Property CcRecipients As List(Of String)

        ''' <summary>
        ''' Gets the recipients that will be placed in the Bcc field.
        ''' </summary>
        <Category("Recipients")>
        <Description("Recipients placed in the Bcc field.")>
        Public ReadOnly Property BccRecipients As List(Of String)

        ''' <summary>
        ''' Gets the full paths of the files that will be attached to the message.
        ''' </summary>
        <Category("Attachments")>
        <Description("Full paths of the files attached to the message.")>
        Public ReadOnly Property Attachments As List(Of String)

        ''' <summary>
        ''' Gets or sets the message subject.
        ''' </summary>
        <Category("Message")>
        <Description("Subject of the e-mail message.")>
        <DefaultValue("")>
        Public Property Subject As String = String.Empty

        ''' <summary>
        ''' Gets or sets the message body.
        ''' </summary>
        <Category("Message")>
        <Description("Body of the e-mail message.")>
        <DefaultValue("")>
        Public Property Body As String = String.Empty

        ''' <summary>
        ''' Gets or sets a value indicating whether <see cref="Body"/> contains HTML.
        ''' </summary>
        <Category("Message")>
        <Description("Indicates whether the body contains HTML.")>
        <DefaultValue(True)>
        Public Property IsBodyHtml As Boolean = True

        ''' <summary>
        ''' Gets or sets a value indicating whether the default Outlook signature should be preserved.
        ''' </summary>
        ''' <remarks>
        ''' Outlook inserts the configured signature when the editor is displayed. When this option is enabled,
        ''' the supplied body is inserted before the generated signature.
        ''' </remarks>
        <Category("Message")>
        <Description("Preserves the default Outlook signature and inserts the supplied body before it.")>
        <DefaultValue(True)>
        Public Property IncludeDefaultSignature As Boolean = True

        ''' <summary>
        ''' Adds a recipient to the To field.
        ''' </summary>
        ''' <param name="address">E-mail address to add.</param>
        ''' <returns>The current message instance.</returns>
        Public Function AddTo(address As String) As OutlookEmailMessage
            AddRecipient(ToRecipients, address)
            Return Me
        End Function

        ''' <summary>
        ''' Adds a recipient to the Cc field.
        ''' </summary>
        ''' <param name="address">E-mail address to add.</param>
        ''' <returns>The current message instance.</returns>
        Public Function AddCc(address As String) As OutlookEmailMessage
            AddRecipient(CcRecipients, address)
            Return Me
        End Function

        ''' <summary>
        ''' Adds a recipient to the Bcc field.
        ''' </summary>
        ''' <param name="address">E-mail address to add.</param>
        ''' <returns>The current message instance.</returns>
        Public Function AddBcc(address As String) As OutlookEmailMessage
            AddRecipient(BccRecipients, address)
            Return Me
        End Function

        ''' <summary>
        ''' Adds a file to the attachment collection.
        ''' </summary>
        ''' <param name="filePath">Full path of the file to attach.</param>
        ''' <returns>The current message instance.</returns>
        Public Function AddAttachment(filePath As String) As OutlookEmailMessage
            If Not String.IsNullOrWhiteSpace(filePath) Then Attachments.Add(filePath)
            Return Me
        End Function

        Private Shared Sub AddRecipient(recipients As ICollection(Of String), address As String)
            If Not String.IsNullOrWhiteSpace(address) Then recipients.Add(address.Trim())
        End Sub
    End Class

End Namespace
