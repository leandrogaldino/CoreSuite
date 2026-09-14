Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>
''' Provides methods for composing e-mail messages in Microsoft Outlook Classic.
''' </summary>
''' <remarks>
''' This class uses Outlook COM automation at runtime and does not require a compile-time reference to
''' Microsoft.Office.Interop.Outlook. It is intended for Microsoft Outlook Classic on Windows.
''' The new Outlook for Windows does not expose the same COM automation model.
''' </remarks>
Public NotInheritable Class OutlookEmail
        Private Const OutlookProgId As String = "Outlook.Application"
        Private Const MailItemType As Integer = 0

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Gets a value indicating whether Microsoft Outlook Classic is available for COM automation.
        ''' </summary>
        Public Shared ReadOnly Property IsAvailable As Boolean
            Get
                Return GetOutlookType() IsNot Nothing
            End Get
        End Property

        ''' <summary>
        ''' Opens the Microsoft Outlook Classic editor with the supplied message data.
        ''' </summary>
        ''' <param name="message">Message to compose.</param>
        ''' <returns>An <see cref="OutlookEmailResult"/> describing the operation result.</returns>
        ''' <remarks>
        ''' This method never sends the e-mail automatically. It only displays the Outlook editor so the user
        ''' can review, edit and send the message manually.
        ''' </remarks>
        Public Shared Function Display(message As OutlookEmailMessage) As OutlookEmailResult
            If message Is Nothing Then
                Return New OutlookEmailResult(OutlookEmailStatus.InvalidMessage, "The e-mail message was not provided.", Nothing)
            End If

            Dim ValidationError = Validate(message)
            If ValidationError IsNot Nothing Then
                Return New OutlookEmailResult(OutlookEmailStatus.InvalidMessage, ValidationError, Nothing)
            End If

            Dim OutlookType = GetOutlookType()
            If OutlookType Is Nothing Then
                Return New OutlookEmailResult(OutlookEmailStatus.OutlookNotAvailable, "Microsoft Outlook Classic is not installed or is not available on this computer.", Nothing)
            End If

            Dim Application As Object = Nothing
            Dim MailItem As Object = Nothing
            Dim AttachmentCollection As Object = Nothing

            Try
                Application = Activator.CreateInstance(OutlookType)
                If Application Is Nothing Then
                    Return New OutlookEmailResult(OutlookEmailStatus.OutlookNotAvailable, "Microsoft Outlook Classic could not be started.", Nothing)
                End If

                MailItem = InvokeMethod(Application, "CreateItem", MailItemType)
                If MailItem Is Nothing Then
                    Return New OutlookEmailResult(OutlookEmailStatus.Failed, "The e-mail message could not be created in Microsoft Outlook.", Nothing)
                End If

                SetProperty(MailItem, "To", JoinRecipients(message.ToRecipients))
                SetProperty(MailItem, "CC", JoinRecipients(message.CcRecipients))
                SetProperty(MailItem, "BCC", JoinRecipients(message.BccRecipients))
                SetProperty(MailItem, "Subject", message.Subject)

                AttachmentCollection = GetProperty(MailItem, "Attachments")
                For Each FilePath In message.Attachments
                    InvokeMethod(AttachmentCollection, "Add", FilePath)
                Next

                InvokeMethod(MailItem, "Display", False)
                ApplyBody(MailItem, message)

                Return New OutlookEmailResult(OutlookEmailStatus.Success, "The message was opened in Microsoft Outlook.", Nothing)
            Catch ex As COMException
                Return New OutlookEmailResult(OutlookEmailStatus.Failed, "The message could not be opened in Microsoft Outlook.", ex)
            Catch ex As Exception
                Return New OutlookEmailResult(OutlookEmailStatus.Failed, "An error occurred while preparing the e-mail message.", ex)
            Finally
                ReleaseComObject(AttachmentCollection)
                ReleaseComObject(MailItem)
                ReleaseComObject(Application)
            End Try
        End Function

        Private Shared Function Validate(message As OutlookEmailMessage) As String
            If message.ToRecipients.Count = 0 AndAlso message.CcRecipients.Count = 0 AndAlso message.BccRecipients.Count = 0 Then
                Return "At least one recipient must be provided."
            End If

            For Each FilePath In message.Attachments
                If String.IsNullOrWhiteSpace(FilePath) Then Return "An attachment has no file path."
                If Not IO.File.Exists(FilePath) Then Return $"The attachment file was not found: {FilePath}"
            Next

            Return Nothing
        End Function

        Private Shared Sub ApplyBody(mailItem As Object, message As OutlookEmailMessage)
            If message.IsBodyHtml Then
                Dim Signature = If(message.IncludeDefaultSignature, Convert.ToString(GetProperty(mailItem, "HTMLBody")), String.Empty)
                SetProperty(mailItem, "HTMLBody", message.Body & Signature)
            Else
                Dim Signature = If(message.IncludeDefaultSignature, Convert.ToString(GetProperty(mailItem, "Body")), String.Empty)

                If message.IncludeDefaultSignature AndAlso Not String.IsNullOrEmpty(message.Body) AndAlso Not String.IsNullOrEmpty(Signature) Then
                    SetProperty(mailItem, "Body", message.Body & Environment.NewLine & Environment.NewLine & Signature)
                Else
                    SetProperty(mailItem, "Body", message.Body & Signature)
                End If
            End If
        End Sub

        Private Shared Function JoinRecipients(recipients As IEnumerable(Of String)) As String
            Return String.Join(";", recipients.Where(Function(Address) Not String.IsNullOrWhiteSpace(Address)))
        End Function

        Private Shared Function GetOutlookType() As Type
            Try
                Return Type.GetTypeFromProgID(OutlookProgId, throwOnError:=False)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function GetProperty(instance As Object, propertyName As String) As Object
            Return instance.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, Nothing, instance, Nothing)
        End Function

        Private Shared Sub SetProperty(instance As Object, propertyName As String, value As Object)
            instance.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, Nothing, instance, New Object() {value})
        End Sub

        Private Shared Function InvokeMethod(instance As Object, methodName As String, ParamArray arguments() As Object) As Object
            Return instance.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, Nothing, instance, arguments)
        End Function

        Private Shared Sub ReleaseComObject(instance As Object)
            If instance Is Nothing OrElse Not Marshal.IsComObject(instance) Then Return

            Try
                Marshal.FinalReleaseComObject(instance)
            Catch
            End Try
        End Sub
End Class