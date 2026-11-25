Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class FrmMessageBox
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2
    Private ReadOnly _Options As CMessageBoxOptions

    Public Sub New(Options As CMessageBoxOptions)
        If Options Is Nothing Then Throw New ArgumentNullException("Options")

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

    Private Sub BtnClose_VisualStateChanged(sender As Object, e As EventArgs) Handles BtnClose.GotFocus, BtnClose.LostFocus, BtnClose.MouseEnter, BtnClose.MouseLeave, BtnClose.MouseDown, BtnClose.MouseUp, BtnClose.KeyDown, BtnClose.KeyUp
        BtnClose.BackColor = If(BtnClose.Focused, BtnClose.FlatAppearance.MouseOverBackColor, TlpTopBar.BackColor)
        BtnClose.Invalidate()
    End Sub

    Private Sub BtnClose_Paint(sender As Object, e As PaintEventArgs) Handles BtnClose.Paint
        Dim Highlighted As Boolean = BtnClose.Focused OrElse BtnClose.ClientRectangle.Contains(BtnClose.PointToClient(Cursor.Position))
        Dim IconColor As Color = If(Highlighted, Color.White, LblTitle.ForeColor)
        Dim X As Single = BtnClose.ClientSize.Width / 2.0F
        Dim Y As Single = BtnClose.ClientSize.Height / 2.0F
        Const Radius As Single = 4.5F
        Dim State = e.Graphics.Save()
        Try
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using Pen As New Pen(IconColor, 1.0F)
                e.Graphics.DrawLine(Pen, X - Radius, Y - Radius, X + Radius, Y + Radius)
                e.Graphics.DrawLine(Pen, X + Radius, Y - Radius, X - Radius, Y + Radius)
            End Using
        Finally
            e.Graphics.Restore(State)
        End Try
    End Sub

    Private Sub PnlMessage_Resize(sender As Object, e As EventArgs) Handles PnlMessage.Resize, PnlIcon.Resize
        AdjustMessagePosition()
    End Sub

    Private Function GetMessageTextHeight() As Integer
        Dim AvailableWidth As Integer = Math.Max(1, PnlMessage.ClientSize.Width - 20)
        LblMessage.MaximumSize = New Size(AvailableWidth, 0)
        Return LblMessage.GetPreferredSize(New Size(AvailableWidth, 0)).Height
    End Function

    Private Sub AdjustMessagePosition()
        Dim TextHeight As Integer = GetMessageTextHeight()
        LblMessage.Location = New Point(10, Math.Max(10, (PnlMessage.ClientSize.Height - TextHeight) \ 2))
        PnlMessage.AutoScrollMinSize = If(TextHeight + 20 > PnlMessage.ClientSize.Height,
                                         New Size(0, TextHeight + 20), Size.Empty)
        AdjustIconAlignment(TextHeight)
    End Sub

    Private Sub AdjustIconAlignment(TextHeight As Integer)
        If PnlIcon Is Nothing OrElse PbxIcon Is Nothing Then Return

        Dim X As Integer = Math.Max(0, (PnlIcon.ClientSize.Width - PbxIcon.Width) \ 2)
        Dim AlignToTop As Boolean = TextHeight > PbxIcon.Height OrElse PnlMessage.AutoScrollMinSize.Height > PnlMessage.ClientSize.Height
        Dim Y As Integer

        If AlignToTop Then
            Y = 10
        Else
            Y = Math.Max(0, (PnlIcon.ClientSize.Height - PbxIcon.Height) \ 2)
        End If

        PbxIcon.Location = New Point(X, Y)
    End Sub


    Friend Sub AdjustDialogSize()
        PerformLayout()
        TlpContainer.PerformLayout()
        TlpBody.PerformLayout()
        Dim TextHeight As Integer = GetMessageTextHeight()
        Dim ErrorHeight As Integer = CInt(TlpBody.RowStyles(0).Height)
        Dim DesiredHeight As Integer = 186 + ErrorHeight + Math.Max(0, TextHeight - 72)
        Dim MaxHeight As Integer = Math.Max(186, Screen.FromPoint(Cursor.Position).WorkingArea.Height - 40)
        ClientSize = New Size(ClientSize.Width, Math.Min(DesiredHeight, MaxHeight))
        PerformLayout()
        AdjustMessagePosition()
    End Sub

    Friend Sub SetTitle(Title As String, MessageType As CMessageBoxType)
        If String.IsNullOrWhiteSpace(Title) Then
            Select Case MessageType
                Case CMessageBoxType.Question : Title = "Confirmação"
                Case CMessageBoxType.Warning : Title = "Atenção"
                Case CMessageBoxType.Error : Title = "Erro"
                Case CMessageBoxType.Success : Title = "Sucesso"
                Case Else : Title = "Informação"
            End Select
        End If
        LblTitle.Text = Title
        Text = Title
    End Sub

    Friend Sub AllocateButtons(MessageType As CMessageBoxType, HasException As Boolean)
        ClearButtons()

        Select Case MessageType
            Case CMessageBoxType.Information, CMessageBoxType.Success, CMessageBoxType.Warning
                Dim OkButton As Button = AddButton("OK", DialogResult.OK, 2, 0)
                AcceptButton = OkButton
                OkButton.Select()

            Case CMessageBoxType.Question
                Dim NoButton As Button = AddButton("Não", DialogResult.No, 1, 1)
                AddButton("Sim", DialogResult.Yes, 2, 0)
                AcceptButton = NoButton
                NoButton.Select()

            Case CMessageBoxType.Error
                If _Options.ShowExceptionDetails AndAlso HasException Then
                    Dim DetailsButton As Button = AddButton("Detalhes", DialogResult.None, 1, 1)
                    CcException.HostControl = DetailsButton
                End If

                Dim OkButton As Button = AddButton("OK", DialogResult.OK, 2, 0)
                AcceptButton = OkButton
                OkButton.Select()
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
            .Size = New Size(100, 32),
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
