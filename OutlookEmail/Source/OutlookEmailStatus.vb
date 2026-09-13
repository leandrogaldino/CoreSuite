Imports System.ComponentModel

Namespace CoreSuite.OutlookEmail

    ''' <summary>
    ''' Specifies the result of an Outlook e-mail composition operation.
    ''' </summary>
    Public Enum OutlookEmailStatus
        ''' <summary>
        ''' The Outlook editor was opened successfully.
        ''' </summary>
        <Description("Success")>
        Success = 0

        ''' <summary>
        ''' Microsoft Outlook Classic is not installed or is not registered for COM automation.
        ''' </summary>
        <Description("Outlook Classic is not available")>
        OutlookNotAvailable = 1

        ''' <summary>
        ''' The message contains invalid data.
        ''' </summary>
        <Description("Invalid message")>
        InvalidMessage = 2

        ''' <summary>
        ''' An unexpected error occurred while creating or displaying the message.
        ''' </summary>
        <Description("Failed")>
        Failed = 3
    End Enum

End Namespace
