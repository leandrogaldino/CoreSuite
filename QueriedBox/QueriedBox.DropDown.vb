Partial Public Class QueriedBox
    <DebuggerStepThrough>
    Private Sub CloseDropDown()
        If DropDownResultsForm IsNot Nothing Then
            DropDownResultsForm.Close()
            DropDownResultsForm = Nothing
        End If
    End Sub
    <DebuggerStepThrough>
    Private Sub Form_Deactivate(sender As Object, e As EventArgs)
        If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
            If UCase(Text) = UCase(DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).Value.ToString) Then
                AutoFreeze()
            End If
        End If
        CloseDropDown()
    End Sub
    Private Sub DropDownResultsForm_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
        DropDownResultsForm.Dispose()
        DropDownResultsForm = Nothing
    End Sub
    Public Function DropDownVisible() As Boolean
        If DropDownResultsForm Is Nothing Then
            Return False
        Else
            If DropDownResultsForm.Visible Then
                Return True
            Else
                Return False
            End If
        End If
    End Function
End Class
