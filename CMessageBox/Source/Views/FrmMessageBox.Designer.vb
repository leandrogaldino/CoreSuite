<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMessageBox
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        PnlBorder = New Panel()
        TlpContainer = New TableLayoutPanel()
        TlpTopBar = New TableLayoutPanel()
        BtnClose = New NoFocusCueButton()
        LblTitle = New Label()
        PnlBottomArea = New Panel()
        TlpBottomBar = New TableLayoutPanel()
        PnlBottomSeparator = New Panel()
        TlpBody = New TableLayoutPanel()
        PbxIcon = New PictureBox()
        LblErrorCode = New Label()
        PnlMessage = New Panel()
        LblMessage = New Label()
        CcException = New ControlContainer()
        PnlBorder.SuspendLayout()
        TlpContainer.SuspendLayout()
        TlpTopBar.SuspendLayout()
        PnlBottomArea.SuspendLayout()
        TlpBody.SuspendLayout()
        CType(PbxIcon, ComponentModel.ISupportInitialize).BeginInit()
        PnlMessage.SuspendLayout()
        SuspendLayout()
        ' 
        ' PnlBorder
        ' 
        PnlBorder.BorderStyle = BorderStyle.FixedSingle
        PnlBorder.Controls.Add(TlpContainer)
        PnlBorder.Dock = DockStyle.Fill
        PnlBorder.Location = New Point(0, 0)
        PnlBorder.Name = "PnlBorder"
        PnlBorder.Size = New Size(430, 240)
        PnlBorder.TabIndex = 0
        ' 
        ' TlpContainer
        ' 
        TlpContainer.ColumnCount = 1
        TlpContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        TlpContainer.Controls.Add(TlpTopBar, 0, 0)
        TlpContainer.Controls.Add(TlpBody, 0, 1)
        TlpContainer.Controls.Add(PnlBottomArea, 0, 2)
        TlpContainer.Dock = DockStyle.Fill
        TlpContainer.Location = New Point(0, 0)
        TlpContainer.Name = "TlpContainer"
        TlpContainer.RowCount = 3
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 35.0F))
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 55.0F))
        TlpContainer.Size = New Size(428, 238)
        TlpContainer.TabIndex = 1
        ' 
        ' TlpTopBar
        ' 
        TlpTopBar.BackColor = Color.FromArgb(248, 249, 250)
        TlpTopBar.ColumnCount = 2
        TlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        TlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 50.0F))
        TlpTopBar.Controls.Add(BtnClose, 1, 0)
        TlpTopBar.Controls.Add(LblTitle, 0, 0)
        TlpTopBar.Dock = DockStyle.Fill
        TlpTopBar.Location = New Point(0, 0)
        TlpTopBar.Margin = New Padding(0)
        TlpTopBar.Name = "TlpTopBar"
        TlpTopBar.RowCount = 1
        TlpTopBar.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        TlpTopBar.Size = New Size(428, 35)
        TlpTopBar.TabIndex = 0
        ' 
        ' BtnClose
        ' 
        BtnClose.BackColor = Color.FromArgb(248, 249, 250)
        BtnClose.DialogResult = DialogResult.Cancel
        BtnClose.Dock = DockStyle.Fill
        BtnClose.FlatAppearance.BorderSize = 0
        BtnClose.FlatAppearance.MouseDownBackColor = Color.FromArgb(200, 15, 30)
        BtnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 17, 35)
        BtnClose.FlatStyle = FlatStyle.Flat
        BtnClose.Image = My.Resources.ImageResources.Close
        BtnClose.Location = New Point(378, 0)
        BtnClose.Margin = New Padding(0)
        BtnClose.Name = "BtnClose"
        BtnClose.Size = New Size(50, 35)
        BtnClose.TabIndex = 0
        BtnClose.UseVisualStyleBackColor = False
        ' 
        ' LblTitle
        ' 
        LblTitle.Dock = DockStyle.Fill
        LblTitle.Font = New Font("Segoe UI Semibold", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        LblTitle.ForeColor = Color.FromArgb(45, 45, 48)
        LblTitle.Location = New Point(3, 0)
        LblTitle.Name = "LblTitle"
        LblTitle.Size = New Size(372, 35)
        LblTitle.TabIndex = 1
        LblTitle.Text = "Erro"
        LblTitle.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' PnlBottomArea
        ' 
        PnlBottomArea.BackColor = Color.White
        PnlBottomArea.Controls.Add(PnlBottomSeparator)
        PnlBottomArea.Controls.Add(TlpBottomBar)
        PnlBottomArea.Dock = DockStyle.Fill
        PnlBottomArea.Location = New Point(0, 183)
        PnlBottomArea.Margin = New Padding(0)
        PnlBottomArea.Name = "PnlBottomArea"
        PnlBottomArea.Size = New Size(428, 55)
        PnlBottomArea.TabIndex = 1
        ' 
        ' TlpBottomBar
        ' 
        TlpBottomBar.BackColor = Color.White
        TlpBottomBar.ColumnCount = 4
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120.0F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120.0F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 8.0F))
        TlpBottomBar.Dock = DockStyle.Fill
        TlpBottomBar.Location = New Point(0, 0)
        TlpBottomBar.Margin = New Padding(0)
        TlpBottomBar.Name = "TlpBottomBar"
        TlpBottomBar.RowCount = 1
        TlpBottomBar.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        TlpBottomBar.Size = New Size(428, 55)
        TlpBottomBar.TabIndex = 1
        ' 
        ' PnlBottomSeparator
        ' 
        PnlBottomSeparator.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        PnlBottomSeparator.BackColor = Color.FromArgb(220, 223, 227)
        PnlBottomSeparator.Location = New Point(16, 0)
        PnlBottomSeparator.Margin = New Padding(0)
        PnlBottomSeparator.Name = "PnlBottomSeparator"
        PnlBottomSeparator.Size = New Size(396, 1)
        PnlBottomSeparator.TabIndex = 0
        ' 
        ' TlpBody
        ' 
        TlpBody.ColumnCount = 2
        TlpBody.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100.0F))
        TlpBody.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        TlpBody.Controls.Add(PbxIcon, 0, 1)
        TlpBody.Controls.Add(LblErrorCode, 0, 0)
        TlpBody.Controls.Add(PnlMessage, 1, 1)
        TlpBody.Dock = DockStyle.Fill
        TlpBody.Location = New Point(3, 38)
        TlpBody.Name = "TlpBody"
        TlpBody.RowCount = 2
        TlpBody.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        TlpBody.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        TlpBody.Size = New Size(422, 142)
        TlpBody.TabIndex = 2
        ' 
        ' PbxIcon
        ' 
        PbxIcon.Dock = DockStyle.Fill
        PbxIcon.Image = My.Resources.ImageResources.Information
        PbxIcon.Location = New Point(3, 33)
        PbxIcon.Name = "PbxIcon"
        PbxIcon.Size = New Size(94, 106)
        PbxIcon.SizeMode = PictureBoxSizeMode.CenterImage
        PbxIcon.TabIndex = 5
        PbxIcon.TabStop = False
        ' 
        ' LblErrorCode
        ' 
        TlpBody.SetColumnSpan(LblErrorCode, 2)
        LblErrorCode.Dock = DockStyle.Fill
        LblErrorCode.Font = New Font("Segoe UI Semibold", 9.75F, FontStyle.Bold)
        LblErrorCode.Location = New Point(3, 3)
        LblErrorCode.Margin = New Padding(3)
        LblErrorCode.Name = "LblErrorCode"
        LblErrorCode.Size = New Size(416, 24)
        LblErrorCode.TabIndex = 4
        LblErrorCode.Text = "ERRO CLN008"
        LblErrorCode.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' PnlMessage
        ' 
        PnlMessage.AutoScroll = True
        PnlMessage.Controls.Add(LblMessage)
        PnlMessage.Dock = DockStyle.Fill
        PnlMessage.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        PnlMessage.Location = New Point(103, 33)
        PnlMessage.Name = "PnlMessage"
        PnlMessage.Padding = New Padding(5)
        PnlMessage.Size = New Size(316, 106)
        PnlMessage.TabIndex = 3
        ' 
        ' LblMessage
        ' 
        LblMessage.AutoSize = True
        LblMessage.BackColor = Color.Transparent
        LblMessage.Location = New Point(10, 10)
        LblMessage.Name = "LblMessage"
        LblMessage.Size = New Size(37, 17)
        LblMessage.TabIndex = 4
        LblMessage.Text = "Body"
        LblMessage.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' CcException
        ' 
        CcException.DropDownBorderColor = SystemColors.HotTrack
        CcException.HostControl = Nothing
        CcException.HostedControl = Nothing
        ' 
        ' FrmMessageBox
        ' 
        AutoScaleMode = AutoScaleMode.None
        BackColor = Color.White
        ClientSize = New Size(430, 240)
        Controls.Add(PnlBorder)
        FormBorderStyle = FormBorderStyle.None
        Name = "FrmMessageBox"
        ShowIcon = False
        StartPosition = FormStartPosition.CenterScreen
        Text = "Frm"
        PnlBorder.ResumeLayout(False)
        TlpContainer.ResumeLayout(False)
        TlpTopBar.ResumeLayout(False)
        PnlBottomArea.ResumeLayout(False)
        TlpBody.ResumeLayout(False)
        CType(PbxIcon, ComponentModel.ISupportInitialize).EndInit()
        PnlMessage.ResumeLayout(False)
        PnlMessage.PerformLayout()
        ResumeLayout(False)
    End Sub

    Friend WithEvents PnlBorder As Panel
    Friend WithEvents TlpContainer As TableLayoutPanel
    Friend WithEvents TlpTopBar As TableLayoutPanel
    Friend WithEvents BtnClose As NoFocusCueButton
    Friend WithEvents LblTitle As Label
    Friend WithEvents PnlBottomArea As Panel
    Friend WithEvents TlpBottomBar As TableLayoutPanel
    Friend WithEvents PnlBottomSeparator As Panel
    Friend WithEvents TlpBody As TableLayoutPanel
    Friend WithEvents PbxIcon As PictureBox
    Friend WithEvents LblErrorCode As Label
    Friend WithEvents LblMessage As Label
    Friend WithEvents PnlMessage As Panel
    Friend WithEvents CcException As ControlContainer
End Class