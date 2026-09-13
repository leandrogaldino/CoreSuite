Imports System.ComponentModel

Namespace CoreSuite.OutlookEmail

    ''' <summary>
    ''' Represents the result of an attempt to open an e-mail message in Microsoft Outlook Classic.
    ''' </summary>
    <DebuggerDisplay("{Status}")>
    Public NotInheritable Class OutlookEmailResult
        Friend Sub New(status As OutlookEmailStatus, message As String, exception As Exception)
            Me.Status = status
            Me.Message = message
            Me.Exception = exception
        End Sub

        ''' <summary>
        ''' Gets the operation status.
        ''' </summary>
        <Category("Result")>
        <Description("Status returned by the Outlook e-mail composition operation.")>
        Public ReadOnly Property Status As OutlookEmailStatus

        ''' <summary>
        ''' Gets a value indicating whether the Outlook editor was opened successfully.
        ''' </summary>
        <Category("Result")>
        <Description("Indicates whether the Outlook editor was opened successfully.")>
        Public ReadOnly Property Success As Boolean
            Get
                Return Status = OutlookEmailStatus.Success
            End Get
        End Property

        ''' <summary>
        ''' Gets a user-readable description of the result.
        ''' </summary>
        <Category("Result")>
        <Description("User-readable description of the result.")>
        Public ReadOnly Property Message As String

        ''' <summary>
        ''' Gets the exception raised during the operation, when available.
        ''' </summary>
        <Browsable(False)>
        Public ReadOnly Property Exception As Exception
    End Class

End Namespace
