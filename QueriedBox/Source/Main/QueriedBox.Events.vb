Partial Public Class QueriedBox
    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        SelectionStart = Text.Length
        If Not _FirstEnter Then
            If Not DesignMode AndAlso Parent IsNot Nothing Then
                AddHandler Parent.FindForm.Deactivate, AddressOf Form_Deactivate
                Timer = New Timer With {
                    .Interval = Search.Interval
                }
            End If
            _FirstEnter = True
        End If
    End Sub
    Protected Overrides Sub OnLeave(e As EventArgs)
        MyBase.OnLeave(e)
        AutoFreezeIfMatched()
        If SelectFromStartOnLeave Then
            SelectionStart = 0
        End If
    End Sub
    <DebuggerStepThrough>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        Dim Chars As Integer
        Dim CharsDif As Integer
        If Not DesignMode AndAlso _KeyDown Then
            If Search.Enabled Then
                If Frozen.FrozenValue <> Nothing And Text <> Frozen.FrozenValue Then
                    AutoUnfreeze()
                End If
                If Parent IsNot Nothing Then
                    If DropDownResultsForm Is Nothing Then
                        DropDownResultsForm = New DropDownResultsPopup(Me)
                        CType(DropDownResultsForm, DropDownResultsPopup).Textbox = Me
                        DropDownResultsForm.Location = Me.Parent.PointToScreen(New Point(Me.Left - DropDown.StretchLeft, Me.Bottom))
                        DropDownResultsForm.Width = DropDown.StretchLeft + Me.Width + If(DropDown.AutoStretchRight, 0, DropDown.StretchRight)
                        DropDownResultsForm.Height = DropDown.StretchDown
                        AddHandler DropDownResultsForm.FormClosed, AddressOf DropDownResultsForm_FormClosed
                        DropDownResultsForm.Show()
                        DropDownVisible()
                    End If
                    Chars = Text.Replace("%", Nothing).Length
                    CharsDif = Search.MinimumChars - Chars
                    If Chars = 0 Then
                        CloseDropDown()
                    Else
                        If Chars < Search.MinimumChars Then
                            DropDownResultsForm.DgvResults.DataSource = Nothing
                            DropDownResultsForm.LblCharsRemaining.Visible = True
                            DropDownResultsForm.DgvResults.Visible = False
                            If CharsDif > 1 Then
                                DropDownResultsForm.LblCharsRemaining.Text = String.Format(Messages.CharsRemainingPluralText, CharsDif)
                            ElseIf CharsDif = 1 Then
                                DropDownResultsForm.LblCharsRemaining.Text = String.Format(Messages.CharsRemainingSingularText, CharsDif)
                            End If
                        ElseIf Chars >= Search.MinimumChars Then
                            Timer.Stop() : Timer.Start()
                            DropDownResultsForm.LblCharsRemaining.Visible = False
                            DropDownResultsForm.DgvResults.Visible = True
                        End If
                    End If
                End If
            End If
        End If
        _KeyDown = False
    End Sub
    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        If Search.Enabled Then
            Frozen.SetUnFrozenColor(ForeColor)
        End If
    End Sub
    Protected Overrides Sub OnPreviewKeyDown(e As PreviewKeyDownEventArgs)
        MyBase.OnPreviewKeyDown(e)
        Dim Row As Long
        If Search.Enabled Then
            If Frozen.AllowHyperlink Then FormatTextBox(e.Control)
            If DropDownResultsForm IsNot Nothing Then
                If DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                    Select Case e.KeyCode
                        Case Is = Keys.Tab
                            AutoFreezeIfMatched()
                        Case Is = Keys.Enter
                            AutoFreeze()
                            Me.Select(TextLength, 0)
                            CloseDropDown()
                        Case Is = Keys.Down
                            Row = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                            If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.Rows.Count > Row + 1 Then
                                DropDownResultsForm.DgvResults.Rows(Row + 1).Selected = True
                                If DropDownResultsForm.DgvResults.SelectedRows(0).Index = DropDownResultsForm.DgvResults.Rows.Count - 1 Then
                                    DropDownResultsForm.DgvResults.FirstDisplayedScrollingRowIndex = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                                    Row += 1
                                Else
                                    Row += 2
                                End If
                                If DropDownResultsForm.DgvResults.Rows(Row).Displayed = False Then
                                    If Row >= 3 Then
                                        DropDownResultsForm.DgvResults.FirstDisplayedScrollingRowIndex = DropDownResultsForm.DgvResults.SelectedRows(0).Index - 2
                                    End If
                                End If
                            End If
                        Case Is = Keys.Up
                            Row = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                            If DropDownResultsForm.DgvResults.Visible = True And Row > 0 Then
                                DropDownResultsForm.DgvResults.Rows(Row - 1).Selected = True
                                If DropDownResultsForm.DgvResults.Rows(Row - 1).Displayed = False Then
                                    DropDownResultsForm.DgvResults.FirstDisplayedScrollingRowIndex = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                                End If
                            End If
                    End Select
                End If
                If e.KeyCode = Keys.Escape Then CloseDropDown()
            End If
        End If
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        _KeyDown = True
        If Search.Enabled Then
            Select Case e.KeyCode
                Case Is = Keys.Enter, Keys.Escape
                    e.SuppressKeyPress = True
                Case Is = Keys.Down
                    e.Handled = True
                Case Is = Keys.Up
                    e.Handled = True
            End Select
        End If
    End Sub
    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)
        If Search.Enabled Then
            FormatTextBox(False)
        End If
    End Sub
    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If Search.Enabled Then
            If _IsHyperlink AndAlso e.Button = MouseButtons.Left Then
                OnHyperlinkClicked()
                FormatTextBox(False)
            End If
        End If
    End Sub
    Protected Overridable Sub OnFrozenPrimaryKeyChanging(OldPrimaryKey As Object, NewPrimaryKey As Object)
        RaiseEvent FrozenPrimaryKeyChanging(Me, New FrozenPrimaryKeyEventArgs(OldPrimaryKey, NewPrimaryKey))
    End Sub
    Protected Overridable Sub OnFrozenPrimaryKeyChanged(OldPrimaryKey As Object, NewPrimaryKey As Object)
        RaiseEvent FrozenPrimaryKeyChanged(Me, New FrozenPrimaryKeyEventArgs(OldPrimaryKey, NewPrimaryKey))
    End Sub
    Protected Overridable Sub OnHyperlinkClicked()
        RaiseEvent HyperlinkClicked(Me, New FrozenPrimaryKeyEventArgs(Frozen.FrozenPrimaryKey))
    End Sub
End Class
