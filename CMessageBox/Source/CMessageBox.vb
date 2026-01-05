Imports System.Drawing
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Windows.Forms
Imports System.Xml

Public Class CMessageBox
    Private Shared _Title As String
    Private Shared _Message As String
    Private Shared _BoxType As CMessageBoxType
    Private Shared _Buttons As CMessageBoxButtons
    Private Shared _Exception As Exception
    Private Shared _ErrorFileName As String
    Public Shared Property SmtpMailServer As New SmtpClient
    Public Shared Property MessageBackColor As Color = Color.White
    Public Shared Property PanelButtonsBackColor As Color = Color.WhiteSmoke
    Public Shared Property MessageFont As Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
    Public Shared Property TitleFont As Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
    Public Shared Property MessageForeColor As Color = SystemColors.ControlText
    Public Shared Property TitleForeColor As Color = SystemColors.ControlText
    Public Shared Property ErrorImage As Image = My.Resources._Error
    Public Shared Property InformationImage As Image = My.Resources.Information
    Public Shared Property QuestionImage As Image = My.Resources.Question
    Public Shared Property DoneImage As Image = My.Resources.Done
    Public Shared Property TipImage As Image = My.Resources.Tip
    Public Shared Property WarningImage As Image = My.Resources.Warning
    Public Shared Property ShowEmailErrorButton As Boolean = True
    Public Shared Property ShowSaveErrorButton As Boolean = True
    Public Shared Property MaximumWidth As Integer = 350
    Public Shared Property AditionalExceptionInformation As New Dictionary(Of String, String)
    Public Shared Property ErrorTempFileLocation As String

    Private Shared Function GetImage(Type As CMessageBoxType) As Image
        Select Case Type
            Case Is = CMessageBoxType.Done
                Return DoneImage
            Case Is = CMessageBoxType.Error
                Return ErrorImage
            Case Is = CMessageBoxType.Information
                Return InformationImage
            Case Is = CMessageBoxType.Question
                Return QuestionImage
            Case Is = CMessageBoxType.Tip
                Return TipImage
            Case Else
                Return WarningImage
        End Select
    End Function

    Private Shared Function SaveXmlError(Location As String) As String
        Dim Xml As New XmlDocument
        Dim Node As XmlNode
        Node = Xml.CreateNode(XmlNodeType.Element, "Erro", Nothing)
        Xml.AppendChild(Node)
        Node = Xml.CreateNode(XmlNodeType.Element, "Data", Nothing)
        Node.InnerText = Today.ToLongDateString
        Xml.SelectSingleNode("Erro").AppendChild(Node)
        If Not String.IsNullOrEmpty(_Title) Then
            Node = Xml.CreateNode(XmlNodeType.Element, "Titulo", Nothing)
            Node.InnerText = _Title
            Xml.SelectSingleNode("Erro").AppendChild(Node)
        End If
        If Not String.IsNullOrEmpty(_Message) Then
            Node = Xml.CreateNode(XmlNodeType.Element, "Mensagem_do_Controle", Nothing)
            Node.InnerText = _Message
            Xml.SelectSingleNode("Erro").AppendChild(Node)
        End If
        Node = Xml.CreateNode(XmlNodeType.Element, "Tipo", Nothing)
        Node.InnerText = _Exception.GetType().ToString()
        Xml.SelectSingleNode("Erro").AppendChild(Node)
        Node = Xml.CreateNode(XmlNodeType.Element, "Mensagem_da_Excecao", Nothing)
        Node.InnerText = _Exception.Message
        Xml.SelectSingleNode("Erro").AppendChild(Node)
        Node = Xml.CreateNode(XmlNodeType.Element, "Rastreamento", Nothing)
        Node.InnerText = _Exception.StackTrace
        Xml.SelectSingleNode("Erro").AppendChild(Node)
        If _Exception.InnerException IsNot Nothing Then
            Node = Xml.CreateNode(XmlNodeType.Element, "Erro_Interno", Nothing)
            Xml.SelectSingleNode("Erro").AppendChild(Node)
            Node = Xml.CreateNode(XmlNodeType.Element, "Tipo", Nothing)
            Node.InnerText = _Exception.InnerException.GetType().ToString()
            Xml.SelectSingleNode("Erro/Erro_Interno").AppendChild(Node)
            Node = Xml.CreateNode(XmlNodeType.Element, "Mensagem_da_Excecao", Nothing)
            Node.InnerText = _Exception.InnerException.Message
            Xml.SelectSingleNode("Erro/Erro_Interno").AppendChild(Node)
            Node = Xml.CreateNode(XmlNodeType.Element, "Rastreamento", Nothing)
            Node.InnerText = _Exception.InnerException.StackTrace
            Xml.SelectSingleNode("Erro/Erro_Interno").AppendChild(Node)
        End If
        If TxtExceptionSteps IsNot Nothing AndAlso Not String.IsNullOrEmpty(TxtExceptionSteps.Text) Then
            Node = Xml.CreateNode(XmlNodeType.Element, "Passos", Nothing)
            Node.InnerText = TxtExceptionSteps.Text
            Xml.SelectSingleNode("Erro").AppendChild(Node)
        End If
        If AditionalExceptionInformation.Count > 0 Then
            Node = Xml.CreateNode(XmlNodeType.Element, "Informacoes_Adicionais", Nothing)
            Xml.SelectSingleNode("Erro").AppendChild(Node)
            For Each a In AditionalExceptionInformation
                Node = Xml.CreateNode(XmlNodeType.Element, a.Key, Nothing)
                Node.InnerText = a.Value
                Xml.SelectSingleNode("Erro/Informacoes_Adicionais").AppendChild(Node)
            Next a
        End If
        Using Sw As New StringWriter
            Using Xw As New XmlTextWriter(Sw)
                Xml.WriteTo(Xw)
                File.WriteAllText(Location, Sw.ToString)
                Return Location
            End Using
        End Using
    End Function

    Public Shared Function Show(Message As String) As DialogResult
        _Title = Nothing
        _Message = Message
        _BoxType = Nothing
        _Buttons = Nothing
        _Exception = Nothing
        Return ShowMessage()
    End Function

    Public Shared Function Show(Message As String, BoxType As CMessageBoxType) As DialogResult
        _Title = Nothing
        _Message = Message
        _BoxType = BoxType
        _Buttons = Nothing
        _Exception = Nothing
        Return ShowMessage()
    End Function

    Public Shared Function Show(Message As String, BoxType As CMessageBoxType, BoxButtons As CMessageBoxButtons) As DialogResult
        _Title = Nothing
        _Message = Message
        _BoxType = BoxType
        _Buttons = BoxButtons
        _Exception = Nothing
        Return ShowMessage()
    End Function

    Public Shared Function Show(Message As String, BoxType As CMessageBoxType, BoxButtons As CMessageBoxButtons, ex As Exception) As DialogResult
        _Title = Nothing
        _Message = Message
        _BoxType = BoxType
        _Buttons = BoxButtons
        _Exception = ex
        Return ShowMessage()
    End Function

    Public Shared Function Show(Title As String, Message As String) As DialogResult
        _Title = Title
        _Message = Message
        _BoxType = Nothing
        _Buttons = Nothing
        _Exception = Nothing
        Return ShowMessage()
    End Function

    Public Shared Function Show(Title As String, Message As String, BoxType As CMessageBoxType) As DialogResult
        _Title = Title
        _Message = Message
        _BoxType = BoxType
        _Buttons = Nothing
        _Exception = Nothing
        Return ShowMessage()
    End Function

    Public Shared Function Show(Title As String, Message As String, BoxType As CMessageBoxType, BoxButtons As CMessageBoxButtons) As DialogResult
        _Title = Title
        _Message = Message
        _BoxType = BoxType
        _Buttons = BoxButtons
        _Exception = Nothing
        Return ShowMessage()
    End Function

    Public Shared Function Show(Title As String, Message As String, BoxType As CMessageBoxType, BoxButtons As CMessageBoxButtons, ex As Exception) As DialogResult
        _Title = Title
        _Message = Message
        _BoxType = BoxType
        _Buttons = BoxButtons
        _Exception = ex
        Return ShowMessage()
    End Function

    Private Shared Sub InitializeForm()
        Form = New Form With {
            .FormBorderStyle = FormBorderStyle.FixedSingle,
            .BackColor = MessageBackColor,
            .Font = MessageFont,
            .MaximizeBox = False,
            .MinimizeBox = False,
            .MinimumSize = New Size(350, 243),
            .Size = New Size(350, 243),
            .Padding = New Padding(0, 12, 0, 0),
            .ShowIcon = False,
            .ShowInTaskbar = False,
            .KeyPreview = True,
            .TopMost = True
        }
        AddHandler Form.FormClosed, AddressOf DeleteTempErrorFile
    End Sub

    Private Shared Sub DeleteTempErrorFile()
        If File.Exists(_ErrorFileName) Then File.Delete(_ErrorFileName)
    End Sub

    Private Shared Sub InitializeImage()
        PnImage = New Panel With {
            .Size = New Size(310, 42),
            .Location = New Point(15, 1),
            .BackgroundImage = GetImage(_BoxType),
            .BackgroundImageLayout = ImageLayout.Zoom
        }
        Form.Controls.Add(PnImage)
    End Sub

    Private Shared Sub InitializeLabels()
        LblMessage = New Label With {
            .Text = _Message,
            .Font = MessageFont,
            .ForeColor = MessageForeColor,
            .AutoSize = True,
            .MaximumSize = New Size(MaximumWidth, 0),
            .BackColor = Color.Transparent
        }
        If _Title <> Nothing Then
            LblTitle = New Label With {
                .Text = _Title,
                .Font = TitleFont,
                .ForeColor = TitleForeColor,
                .AutoSize = True,
                .MaximumSize = New Size(MaximumWidth, 0),
                .Location = New Point(15, 55),
                .BackColor = Color.Transparent
            }
            LblMessage.Location = New Point(12, 97)
            Form.Controls.Add(LblTitle)
        Else
            Form.MinimumSize = New Size(350, 230)
            Form.Size = New Size(350, 230)
            LblMessage.Location = New Point(15, 75)
        End If
        Form.Controls.Add(LblMessage)
    End Sub

    Private Shared Sub InitializeButtons()
        Dim LeftButtonLocation As New Point(85, 9)
        Dim MiddleButtonLocation As New Point(166, 9)
        Dim RightButtonLocation As New Point(247, 9)
        PnButtons = New Panel With {
            .BackColor = PanelButtonsBackColor,
            .Dock = DockStyle.Bottom,
            .Height = 50
        }
        Form.Controls.Add(PnButtons)
        BtnAbort = New Button With {
            .DialogResult = DialogResult.Abort,
            .Size = New Size(75, 30),
            .TabIndex = 2,
            .Text = "Abortar",
            .UseVisualStyleBackColor = True,
            .Parent = PnButtons,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        }
        BtnRetry = New Button With {
            .DialogResult = DialogResult.Retry,
            .Size = New Size(75, 30),
            .TabIndex = 0,
            .Text = "Repetir",
            .UseVisualStyleBackColor = True,
            .Parent = PnButtons,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        }
        BtnIgnore = New Button With {
            .DialogResult = DialogResult.Ignore,
            .Size = New Size(75, 30),
            .Text = "Ignorar",
            .TabIndex = 1,
            .UseVisualStyleBackColor = True,
            .Parent = PnButtons,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        }
        BtnOK = New Button With {
            .DialogResult = DialogResult.OK,
            .Size = New Size(75, 30),
            .Text = "OK",
            .TabIndex = 0,
            .UseVisualStyleBackColor = True,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        }
        BtnCancel = New Button With {
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .DialogResult = DialogResult.Cancel,
            .Size = New Size(75, 30),
            .Text = "Cancelar",
            .TabIndex = 1,
            .UseVisualStyleBackColor = True
        }
        BtnYes = New Button With {
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .DialogResult = DialogResult.Yes,
            .Size = New Size(75, 30),
            .TabIndex = 0,
            .Text = "Sim",
            .UseVisualStyleBackColor = True
        }
        BtnNo = New Button With {
            .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
            .DialogResult = DialogResult.No,
            .Size = New Size(75, 30),
            .Text = "Não",
            .TabIndex = 1,
            .UseVisualStyleBackColor = True
        }
        Select Case _Buttons
            Case Is = CMessageBoxButtons.AbortRetryIgnore
                BtnAbort.Location = LeftButtonLocation
                BtnRetry.Location = MiddleButtonLocation
                BtnIgnore.Location = RightButtonLocation
                PnButtons.Controls.AddRange({BtnAbort, BtnRetry, BtnIgnore})
                BtnOK.Dispose()
                BtnCancel.Dispose()
                BtnYes.Dispose()
                BtnNo.Dispose()
            Case Is = CMessageBoxButtons.OK
                BtnOK.Location = RightButtonLocation
                PnButtons.Controls.Add(BtnOK)
                BtnAbort.Dispose()
                BtnRetry.Dispose()
                BtnIgnore.Dispose()
                BtnCancel.Dispose()
                BtnYes.Dispose()
                BtnNo.Dispose()
            Case Is = CMessageBoxButtons.OKCancel
                BtnOK.Location = MiddleButtonLocation
                BtnCancel.Location = RightButtonLocation
                PnButtons.Controls.AddRange({BtnOK, BtnCancel})
                BtnAbort.Dispose()
                BtnRetry.Dispose()
                BtnIgnore.Dispose()
                BtnYes.Dispose()
                BtnNo.Dispose()
            Case Is = CMessageBoxButtons.RetryCancel
                BtnRetry.Location = MiddleButtonLocation
                BtnCancel.Location = RightButtonLocation
                PnButtons.Controls.AddRange({BtnRetry, BtnCancel})
                BtnAbort.Dispose()
                BtnIgnore.Dispose()
                BtnOK.Dispose()
                BtnYes.Dispose()
                BtnNo.Dispose()
            Case Is = CMessageBoxButtons.YesNo
                BtnYes.Location = MiddleButtonLocation
                BtnNo.Location = RightButtonLocation
                PnButtons.Controls.AddRange({BtnYes, BtnNo})
                BtnAbort.Dispose()
                BtnRetry.Dispose()
                BtnIgnore.Dispose()
                BtnOK.Dispose()
                BtnCancel.Dispose()
            Case Is = CMessageBoxButtons.YesNoCancel
                BtnYes.Location = LeftButtonLocation
                BtnNo.Location = MiddleButtonLocation
                BtnCancel.Location = RightButtonLocation
                PnButtons.Controls.AddRange({BtnYes, BtnNo, BtnCancel})
                BtnAbort.Dispose()
                BtnRetry.Dispose()
                BtnIgnore.Dispose()
                BtnOK.Dispose()
            Case Else
                BtnOK.Location = RightButtonLocation
                PnButtons.Controls.Add(BtnOK)
                BtnAbort.Dispose()
                BtnRetry.Dispose()
                BtnIgnore.Dispose()
                BtnCancel.Dispose()
                BtnYes.Dispose()
                BtnNo.Dispose()
        End Select
    End Sub

    Private Shared Sub InitializeExceptionHandler()
        If _BoxType = CMessageBoxType.Error Then
            LblMessage.Cursor = Cursors.Hand
            UcExceptionDialog = New UserControl With {
                .BackColor = Color.White,
                .Font = New Font("Microsoft Sans Serif", 9.75, FontStyle.Regular),
                .Margin = New Padding(4, 4, 4, 4),
                .Size = New Size(700, 400)
            }
            TcException = New TabControl With {
                .Dock = DockStyle.Fill
            }
            TpExceptionDetail = New TabPage With {
                .Text = "Detalhes",
                .UseVisualStyleBackColor = True
            }
            TpExceptionSupport = New TabPage With {
                .Text = "Suporte",
                .UseVisualStyleBackColor = True
            }
            WbDetailException = New WebBrowser With {
                .Dock = DockStyle.Fill,
                .Url = New Uri(SaveXmlError(_ErrorFileName))
            }
            LblExceptionSteps = New Label With {
                .AutoSize = False,
                .Dock = DockStyle.Top,
                .Location = New Point(3, 58),
                .Text = "Escreva aqui em detalhes, os passos que levaram a esse erro",
                .TextAlign = ContentAlignment.MiddleCenter,
                .ForeColor = Color.FromArgb(40, 40, 40)
            }
            TxtExceptionSteps = New TextBox With {
                .Location = New Point(6, 77),
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .Size = New Size(283, 142),
                .TabIndex = 2
            }
            PnExceptionButtons = New Panel With {
                .BackColor = Color.WhiteSmoke,
                .Dock = DockStyle.Bottom,
                .Height = 38,
                .Margin = New Padding(4, 4, 4, 4)
            }
            If ShowSaveErrorButton Then
                BtnSaveException = New Button With {
                    .DialogResult = DialogResult.Abort,
                    .Size = New Size(75, 30),
                    .Left = UcExceptionDialog.Width - 85 * 2,
                    .Top = 4,
                    .Text = "Salvar",
                    .UseVisualStyleBackColor = True,
                    .TabIndex = 3
                }
            End If
            If ShowEmailErrorButton Then
                BtnSendEmailException = New Button With {
                    .DialogResult = DialogResult.Abort,
                    .Size = New Size(75, 30),
                    .Left = UcExceptionDialog.Width - 85,
                    .Top = 4,
                    .Text = "E-Mail",
                    .UseVisualStyleBackColor = True,
                    .TabIndex = 4
                }
            Else
                If BtnSaveException IsNot Nothing Then
                    BtnSaveException.Location = New Point(218, 4)
                End If
            End If
            TpExceptionDetail.Controls.Add(WbDetailException)
            TpExceptionSupport.Controls.AddRange({LblExceptionSteps, TxtExceptionSteps})
            TcException.TabPages.AddRange({TpExceptionDetail, TpExceptionSupport})
            PnExceptionButtons.Controls.AddRange({BtnSaveException, BtnSendEmailException})
            UcExceptionDialog.Controls.AddRange({TcException, PnExceptionButtons})
            CcContainer = New ControlContainer With {
                .DropDownControl = UcExceptionDialog,
                .HostControl = LblMessage
            }
            TxtExceptionSteps.BringToFront()
        End If
    End Sub

    Private Shared Function ShowMessage() As DialogResult
        _ErrorFileName = If(String.IsNullOrEmpty(ErrorTempFileLocation), Path.Combine(Path.GetTempPath(), Guid.NewGuid.ToString() & ".xml"), Path.Combine(ErrorTempFileLocation, Guid.NewGuid.ToString() & ".xml"))
        InitializeForm()
        InitializeImage()
        InitializeButtons()
        InitializeLabels()
        If _Exception IsNot Nothing Then
            InitializeExceptionHandler()
        End If
        Return Form.ShowDialog
    End Function

    Private Shared Sub KeyDown(sender As Object, e As KeyEventArgs) Handles Form.KeyDown
        If e.KeyCode = Keys.Escape Then
            Form.Dispose()
        End If
    End Sub

    Private Shared Sub UcBtnEmail_Click(sender As Object, e As EventArgs) Handles BtnSendEmailException.Click
        Dim Mail As New MailMessage
        Dim Credential As Net.NetworkCredential
        Try
            CcContainer.DropDownControl.Cursor = Cursors.WaitCursor
            Credential = SmtpMailServer.Credentials
            Mail.Subject = "Relatório de Erro do CMessageBox"
            Mail.From = New MailAddress(Credential.UserName)
            Mail.To.Add(Credential.UserName)
            Mail.Body = "Por favor, verifique o arquivo xml anexo."
            Mail.Attachments.Add(New Attachment(_ErrorFileName, MediaTypeNames.Text.Xml))
            SmtpMailServer.Send(Mail)
            CcContainer.DropDownControl.Cursor = Cursors.Default
            MsgBox("E-Mail Enviado.")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            CcContainer.DropDownControl.Cursor = Cursors.Default
            Mail.Dispose()
        End Try
    End Sub

    Private Shared Sub UcBtnSave_Click(sender As Object, e As EventArgs) Handles BtnSaveException.Click
        SfdSave = New SaveFileDialog With {
            .Filter = "Arquivos XML (*.xml)|*.xml",
            .RestoreDirectory = True,
            .Title = "Salvar Erro",
            .FileName = "Erro"
        }
        If SfdSave.ShowDialog() = DialogResult.OK Then
            SaveXmlError(SfdSave.FileName)
        End If
    End Sub

    Private Shared Sub LblMessage_LblTitle_SizeChanged(sender As Object, e As EventArgs) Handles LblMessage.SizeChanged, LblTitle.SizeChanged
        If LblTitle IsNot Nothing Then
            If LblMessage.Text.Length > LblTitle.Text.Length Then
                Form.Width = LblMessage.Width + 40
            Else
                Form.Width = LblTitle.Width + 40
            End If
        Else
            Form.Width = LblMessage.Width + 40
        End If
        Form.Height = LblMessage.Height + 190 + If(LblTitle IsNot Nothing, LblTitle.Height, 0)
    End Sub

    Private Shared Sub LblTitle_SizeChanged(sender As Object, e As EventArgs) Handles LblTitle.SizeChanged
        If LblTitle IsNot Nothing Then LblMessage.Top = LblTitle.Top + LblTitle.Height + 15
    End Sub

    Friend Shared WithEvents SfdSave As SaveFileDialog
    Friend Shared WithEvents TtTips As ToolTip
    Friend Shared WithEvents CcContainer As ControlContainer
    Friend Shared WithEvents Form As Form
    Friend Shared WithEvents PnImage As Panel
    Friend Shared WithEvents LblTitle As Label
    Friend Shared WithEvents LblMessage As Label
    Friend Shared WithEvents PnButtons As Panel
    Friend Shared WithEvents BtnAbort As Button
    Friend Shared WithEvents BtnRetry As Button
    Friend Shared WithEvents BtnIgnore As Button
    Friend Shared WithEvents BtnOK As Button
    Friend Shared WithEvents BtnCancel As Button
    Friend Shared WithEvents BtnYes As Button
    Friend Shared WithEvents BtnNo As Button
    Friend Shared WithEvents UcExceptionDialog As UserControl
    Friend Shared WithEvents TcException As TabControl
    Friend Shared WithEvents TpExceptionDetail As TabPage
    Friend Shared WithEvents TpExceptionSupport As TabPage
    Friend Shared WithEvents LblExceptionSteps As Label
    Friend Shared WithEvents TxtExceptionSteps As TextBox
    Friend Shared WithEvents PnExceptionButtons As Panel
    Friend Shared WithEvents BtnSaveException As Button
    Friend Shared WithEvents BtnSendEmailException As Button
    Friend Shared WithEvents WbDetailException As WebBrowser
End Class