Imports System.Runtime.InteropServices

Public Class FrmMessageBox
    <DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As IntPtr
    End Function

    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2


    Private Sub DragForm(sender As Object, e As MouseEventArgs) Handles TlpTopBar.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub

    Private Sub AdjustMessagePosition()
        LblMessage.MaximumSize = New Size(PnMessage.Width - 10, 0)
        Dim TextHeight As Integer = LblMessage.PreferredHeight
        If TextHeight < PnMessage.Height Then
            LblMessage.Top = (PnMessage.Height - TextHeight) \ 2
            LblMessage.Left = 5
        Else
            LblMessage.Top = 0
            LblMessage.Left = 5
        End If
    End Sub

    Private Sub PanelMessage_Resize(sender As Object, e As EventArgs) Handles PnMessage.Resize
        AdjustMessagePosition()
    End Sub



    Friend Sub AllocateButtons(MessageType As CMessageBoxType)
        Dim Button As Button
        Select Case MessageType
            Case CMessageBoxType.Information, CMessageBoxType.Success, CMessageBoxType.Tip, CMessageBoxType.Warning
                Button = CreateButton()
                Button.Text = "OK"
                Button.DialogResult = DialogResult.OK
                Button.TabIndex = 0
                TlpBottomBar.Controls.Add(Button, 4, 1)
            Case CMessageBoxType.Question
                Button = CreateButton()
                Button.Text = "Não"
                Button.DialogResult = DialogResult.No
                Button.TabIndex = 0
                TlpBottomBar.Controls.Add(Button, 4, 1)
                Button = CreateButton()
                Button.Text = "Sim"
                Button.DialogResult = DialogResult.Yes
                Button.TabIndex = 1
                TlpBottomBar.Controls.Add(Button, 3, 1)
            Case CMessageBoxType.Error
                Button = CreateButton()
                Button.Text = "OK"
                Button.DialogResult = DialogResult.OK
                Button.TabIndex = 0
                TlpBottomBar.Controls.Add(Button, 4, 1)

        End Select
    End Sub


    Private Function CreateButton() As Button
        Dim Btn As New NoFocusCueButton.NoFocusCueButton With {
            .UseVisualStyleBackColor = True,
            .Anchor = AnchorStyles.None,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 9.75F),
            .Margin = New Padding(3, 6, 3, 3),
            .Size = New Size(94, 35),
            .TextAlign = ContentAlignment.MiddleCenter,
                    .FlatStyle = FlatStyle.Flat
        }
        Btn.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180)
        Btn.FlatAppearance.BorderSize = 1
        Btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 225, 225)
        Btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(205, 205, 205)
        Return Btn
    End Function
End Class


Public Enum CMessageBoxType
    Information
    Success
    [Error]
    Warning
    Question
    Tip
End Enum

