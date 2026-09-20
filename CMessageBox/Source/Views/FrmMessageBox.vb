Imports System.Runtime.InteropServices

Public Class FrmMessageBox
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2
    Private ReadOnly _Options As CMessageBoxOptions

    Public Sub New(Options As CMessageBoxOptions)
        ArgumentNullException.ThrowIfNull(Options)

        InitializeComponent()

        _Options = Options

        LblTitle.Font = _Options.TitleFont
        LblMessage.Font = _Options.MessageFont
        LblTitle.ForeColor = _Options.TitleForeColor
        LblMessage.ForeColor = _Options.MessageForeColor
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As IntPtr
    End Function

    Private Sub DragForm(sender As Object, e As MouseEventArgs) Handles TlpTopBar.MouseDown, LblTitle.MouseDown
        If e.Button <> MouseButtons.Left Then Return

        ReleaseCapture()
        SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End Sub

    Private Sub PnlMessage_Resize(sender As Object, e As EventArgs) Handles PnlMessage.Resize
        AdjustMessagePosition()
    End Sub

    Private Sub AdjustMessagePosition()
        LblMessage.MaximumSize = New Size(Math.Max(0, PnlMessage.ClientSize.Width - 20), 0)

        Dim TextHeight As Integer = LblMessage.PreferredHeight

        LblMessage.Left = 10
        LblMessage.Top = If(TextHeight < PnlMessage.ClientSize.Height, (PnlMessage.ClientSize.Height - TextHeight) \ 2, 10)
    End Sub


    Friend Sub AllocateButtons(MessageType As CMessageBoxType)
        ClearButtons()

        Select Case MessageType
            Case CMessageBoxType.Information, CMessageBoxType.Success, CMessageBoxType.Warning
                AddButton("OK", DialogResult.OK, 2, 0)

            Case CMessageBoxType.Question
                AddButton("Não", DialogResult.No, 1, 1)
                AddButton("Sim", DialogResult.Yes, 2, 0)

            Case CMessageBoxType.Error
                If _Options.ShowExceptionDetails Then
                    Dim DetailsButton As Button = AddButton("Detalhes", DialogResult.None, 1, 1)
                    CcException.HostControl = DetailsButton
                End If

                AddButton("OK", DialogResult.OK, 2, 0)
        End Select
    End Sub
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

    Friend Sub SetErrorCode(ErrorCode As String)
        If String.IsNullOrWhiteSpace(ErrorCode) Then
            LblErrorCode.Visible = False
            TlpBody.RowStyles(0).Height = 0
        Else
            LblErrorCode.Text = ErrorCode
            LblErrorCode.Visible = True
            TlpBody.RowStyles(0).Height = 30
        End If
    End Sub

    Private Function AddButton(Text As String, Result As DialogResult, Column As Integer, TabIndex As Integer) As Button
        Dim Button As Button = CreateButton()

        Button.Text = Text
        Button.DialogResult = Result
        Button.TabIndex = TabIndex

        TlpBottomBar.Controls.Add(Button, Column, 0)

        Return Button
    End Function

    Private Function CreateButton() As Button
        Dim Button As New NoFocusCueButton With {
            .UseVisualStyleBackColor = False,
            .Anchor = AnchorStyles.Left,
            .BackColor = Color.FromArgb(248, 248, 248),
            .ForeColor = Color.FromArgb(45, 45, 48),
            .Font = New Font("Segoe UI", 9.75F),
            .Margin = New Padding(3),
            .Size = New Size(110, 34),
            .TextAlign = ContentAlignment.MiddleCenter,
            .FlatStyle = FlatStyle.Flat
        }

        Button.FlatAppearance.BorderColor = Color.FromArgb(190, 193, 197)
        Button.FlatAppearance.BorderSize = 1
        Button.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 228, 232)
        Button.FlatAppearance.MouseDownBackColor = Color.FromArgb(205, 209, 214)

        Return Button
    End Function

    Private Sub ClearButtons()
        For Each Button In TlpBottomBar.Controls.OfType(Of Button).ToArray()
            TlpBottomBar.Controls.Remove(Button)
            Button.Dispose()
        Next
    End Sub
End Class