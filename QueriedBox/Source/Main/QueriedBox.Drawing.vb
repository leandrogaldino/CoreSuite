
Partial Public Class QueriedBox
    ''' <summary>
    ''' Applies the visual formatting required to display the control as a hyperlink or standard text box.
    ''' </summary>
    ''' <param name="ShowAsLink">
    ''' Indicates whether the control should be formatted as a hyperlink.
    ''' </param>
    Private Sub FormatTextBox(ByVal ShowAsLink As Boolean)
        If Not _CtrlHyperlink Then Return
        If ShowAsLink Then
            Font = New Font(Font, FontStyle.Underline)
            Cursor = Cursors.Hand
            _IsHyperlink = True
        Else
            Font = New Font(Font, FontStyle.Regular)
            Cursor = Cursors.IBeam
            _IsHyperlink = False
        End If
    End Sub
End Class