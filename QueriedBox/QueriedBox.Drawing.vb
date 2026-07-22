Partial Public Class QueriedBox
    Private Sub FormatTextBox(ByVal ShowAsLink As Boolean)
        If Not _CtrlHyperLink Then Return
        If ShowAsLink Then
            Font = New Font(Font, FontStyle.Underline)
            Cursor = Cursors.Hand
            _IsHyperLink = True
        Else
            Font = New Font(Font, FontStyle.Regular)
            Cursor = Cursors.IBeam
            _IsHyperLink = False
        End If
    End Sub
End Class
