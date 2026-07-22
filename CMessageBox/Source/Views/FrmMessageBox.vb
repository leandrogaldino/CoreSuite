Imports System.Runtime.InteropServices

Public Class FrmMessageBox
    Private _Options As CMessageBoxOptions
    Private _Container As UcException
    Public Sub New(Options As CMessageBoxOptions)
        InitializeComponent()
        _Options = Options
        LblTitle.Font = _Options.TitleFont
        LblMessage.Font = _Options.MessageFont
        LblTitle.ForeColor = _Options.TitleForeColor
        LblMessage.ForeColor = _Options.MessageForeColor
        If _Options.ShowExceptionDetails Then
            _Container = New UcException
            CcException.DropDownControl = _Container
        End If
    End Sub
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
            Case CMessageBoxType.Information, CMessageBoxType.Success, CMessageBoxType.Warning
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
                Button.TabIndex = 1
                TlpBottomBar.Controls.Add(Button, 4, 1)
                If _Options.ShowExceptionDetails Then
                    Button = CreateButton()
                    Button.Text = "Detalhes"
                    Button.TabIndex = 0
                    TlpBottomBar.Controls.Add(Button, 3, 1)
                    CcException.HostControl = Button
                End If
        End Select
    End Sub
    Private Function CreateButton() As Button
        Dim Btn As New NoFocusCueButton With {
            .UseVisualStyleBackColor = True,
            .Anchor = AnchorStyles.None,
            .BackColor = Color.White,
            .Font = New Font("Segoe UI", 9.75F),
            .Margin = New Padding(3, 6, 3, 3),
            .Size = New Size(94, 35),
            .TextAlign = ContentAlignment.MiddleCenter,
                    .FlatStyle = FlatStyle.Flat
        }
        Btn.FlatAppearance.BorderColor = Color.Gainsboro
        Btn.FlatAppearance.BorderSize = 1
        Btn.FlatAppearance.MouseOverBackColor = Color.LightGray
        Btn.FlatAppearance.MouseDownBackColor = Color.Silver
        Return Btn
    End Function
    Friend Sub SetMessageIcon(MessageType As CMessageBoxType)
        Select Case MessageType
            Case CMessageBoxType.Error
                PbxIcon.Image = _Options.ErrorImage
            Case CMessageBoxType.Question
                PbxIcon.Image = _Options.QuestionImage
            Case CMessageBoxType.Success
                PbxIcon.Image = _Options.SuccessImage
            Case CMessageBoxType.Warning
                PbxIcon.Image = _Options.WarningImage
            Case Else
                PbxIcon.Image = _Options.InformationImage
        End Select
    End Sub

End Class




