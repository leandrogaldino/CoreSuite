Imports System.Diagnostics
''' <summary>
''' Provides a message filter that detects mouse interactions outside a popup window and closes it when necessary.
''' </summary>
Friend Class PopupWindowHelperMessageFilter
    Implements IMessageFilter
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const WM_NCRBUTTONDOWN As Integer = &HA4
    Private Const WM_NCMBUTTONDOWN As Integer = &HA7
    Private ReadOnly TextBox As Control = Nothing
    ''' <summary>
    ''' Gets or sets the popup window monitored by this message filter.
    ''' </summary>
    Public Property Popup As Form = Nothing
    ''' <summary>
    ''' Initializes a new instance of the <see cref="PopupWindowHelperMessageFilter"/> class.
    ''' </summary>
    ''' <param name="popupW">
    ''' The popup window to monitor.
    ''' </param>
    ''' <param name="textbox">
    ''' The control associated with the popup.
    ''' </param>
    Public Sub New(popupW As Form, textbox As Control)
        Popup = popupW
        Me.TextBox = textbox
    End Sub
    ''' <summary>
    ''' Handles mouse button interactions and closes the popup when the click occurs outside the popup and associated control.
    ''' </summary>
    <DebuggerStepThrough>
    Private Sub OnMouseDown()
        Dim CursorPos As Point = Cursor.Position
        Dim Control = TryCast(TextBox, QueriedBox)
        If Control Is Nothing OrElse TextBox.Parent Is Nothing Then Exit Sub
        If Not Popup.Bounds.Contains(CursorPos) AndAlso
       Not TextBox.Bounds.Contains(TextBox.Parent.PointToClient(CursorPos)) Then
            Control.AutoFreezeIfMatched()
            Popup.Close()
        End If
    End Sub
    ''' <summary>
    ''' Filters Windows messages and handles mouse button events used to determine whether the popup should be closed.
    ''' </summary>
    ''' <param name="m">
    ''' The Windows message being processed.
    ''' </param>
    ''' <returns>
    ''' Returns <see langword="false"/> to allow the message to continue processing.
    ''' </returns>
    <DebuggerStepThrough>
    Private Function IMessageFilter_PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
        If Popup IsNot Nothing Then
            Select Case m.Msg
                Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
                    OnMouseDown()
            End Select
        End If
        Return False
    End Function
End Class