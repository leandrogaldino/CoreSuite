Friend Class PopupWindowHelperMessageFilter
    Implements IMessageFilter
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const WM_NCRBUTTONDOWN As Integer = &HA4
    Private Const WM_NCMBUTTONDOWN As Integer = &HA7
    Private ReadOnly TextBox As Control = Nothing
    Public Property Popup As Form = Nothing
    Public Sub New(ByVal popupW As Form, ByVal textbox As Control)
        Popup = popupW
        Me.TextBox = textbox
    End Sub
    <DebuggerStepThrough>
    Private Sub OnMouseDown()
        Dim cursorPos As Point = Cursor.Position
        If TextBox.Parent Is Nothing Then Exit Sub
        If Not Popup.Bounds.Contains(cursorPos) Then
            If Not TextBox.Bounds.Contains(TextBox.Parent.PointToClient(cursorPos)) Then
                If CType(TextBox, QueriedBox).DropDownResultsForm IsNot Nothing AndAlso CType(TextBox, QueriedBox).DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                    If UCase(CType(TextBox, QueriedBox).Text) = UCase(CType(TextBox, QueriedBox).DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(CType(TextBox, QueriedBox).DisplayFieldAlias = Nothing, CType(TextBox, QueriedBox).DisplayFieldName, CType(TextBox, QueriedBox).DisplayFieldAlias)).Value.ToString) Then
                        CType(TextBox, QueriedBox).AutoFreeze()
                    End If
                End If
                Application.RemoveMessageFilter(Me)
                Popup.Close()
            End If
        End If
    End Sub
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
