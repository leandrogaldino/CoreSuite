Partial Public Class AsyncLookupBox
    Private Sub InitializeSelectionAppearance()
        _UnselectedBackColor = MyBase.BackColor
        _UnselectedForeColor = MyBase.ForeColor
        ApplySelectionAppearance()
    End Sub
    Private Sub ApplySelectionAppearance()
        If _UpdatingSelectionAppearance Then Return
        _UpdatingSelectionAppearance = True
        Try
            If HasSelection AndAlso HighlightSelectedItem Then
                MyBase.BackColor = SelectedItemBackColor
                MyBase.ForeColor = SelectedItemForeColor
            Else
                If Not _UnselectedBackColor.IsEmpty Then MyBase.BackColor = _UnselectedBackColor
                If Not _UnselectedForeColor.IsEmpty Then MyBase.ForeColor = _UnselectedForeColor
            End If
        Finally
            _UpdatingSelectionAppearance = False
        End Try
        UpdateActionButtonLayout()
    End Sub
End Class
