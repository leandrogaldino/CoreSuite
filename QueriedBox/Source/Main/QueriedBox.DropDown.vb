Partial Public Class QueriedBox
    ''' <summary>
    ''' Closes the query results dropdown window and releases its reference.
    ''' </summary>
    <DebuggerStepThrough>
    Private Sub CloseDropDown()
        If DropDownResultsForm IsNot Nothing Then
            DropDownResultsForm.Close()
            DropDownResultsForm = Nothing
        End If
    End Sub
    ''' <summary>
    ''' Handles the control deactivation event and applies automatic freezing behavior when applicable.
    ''' </summary>
    ''' <param name="sender">
    ''' The object that raised the event.
    ''' </param>
    ''' <param name="e">
    ''' The event data.
    ''' </param>
    <DebuggerStepThrough>
    Private Sub Form_Deactivate(sender As Object, e As EventArgs)
        AutoFreezeIfMatched()
    End Sub
    ''' <summary>
    ''' Handles the dropdown form closed event and releases the associated resources.
    ''' </summary>
    ''' <param name="sender">
    ''' The object that raised the event.
    ''' </param>
    ''' <param name="e">
    ''' The event data.
    ''' </param>
    Private Sub DropDownResultsForm_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
        DropDownResultsForm.Dispose()
        DropDownResultsForm = Nothing
    End Sub
    ''' <summary>
    ''' Gets a value indicating whether the query results dropdown is currently visible.
    ''' </summary>
    ''' <returns>
    ''' <see langword="true"/> if the dropdown window is visible; otherwise, <see langword="false"/>.
    ''' </returns>
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