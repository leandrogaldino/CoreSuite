<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMessageBox
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        TlpContainer = New TableLayoutPanel()
        TlpTopBar = New TableLayoutPanel()
        BtnClose = New NoFocusCueButton()
        TlpBottomBar = New TableLayoutPanel()
        TlpBody = New TableLayoutPanel()
        PbxIcon = New PictureBox()
        LblTitle = New Label()
        PnMessage = New Panel()
        LblMessage = New Label()
        CcException = New ControlContainer()
        TlpContainer.SuspendLayout()
        TlpTopBar.SuspendLayout()
        TlpBody.SuspendLayout()
        CType(PbxIcon, ComponentModel.ISupportInitialize).BeginInit()
        PnMessage.SuspendLayout()
        SuspendLayout()
        ' 
        ' TlpContainer
        ' 
        TlpContainer.ColumnCount = 1
        TlpContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TlpContainer.Controls.Add(TlpTopBar, 0, 0)
        TlpContainer.Controls.Add(TlpBottomBar, 0, 2)
        TlpContainer.Controls.Add(TlpBody, 0, 1)
        TlpContainer.Dock = DockStyle.Fill
        TlpContainer.Location = New Point(0, 0)
        TlpContainer.Name = "TlpContainer"
        TlpContainer.RowCount = 3
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 40F))
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 50F))
        TlpContainer.Size = New Size(467, 258)
        TlpContainer.TabIndex = 0
        ' 
        ' TlpTopBar
        ' 
        TlpTopBar.BackColor = SystemColors.Control
        TlpTopBar.ColumnCount = 2
        TlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TlpTopBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 50F))
        TlpTopBar.Controls.Add(BtnClose, 1, 0)
        TlpTopBar.Dock = DockStyle.Fill
        TlpTopBar.Location = New Point(3, 3)
        TlpTopBar.Name = "TlpTopBar"
        TlpTopBar.RowCount = 1
        TlpTopBar.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TlpTopBar.Size = New Size(461, 34)
        TlpTopBar.TabIndex = 0
        ' 
        ' BtnClose
        ' 
        BtnClose.DialogResult = DialogResult.Cancel
        BtnClose.Dock = DockStyle.Fill
        BtnClose.FlatAppearance.BorderSize = 0
        BtnClose.FlatAppearance.MouseOverBackColor = Color.LightCoral
        BtnClose.FlatStyle = FlatStyle.Flat
        BtnClose.Image = My.Resources.ImageResources.Close
        BtnClose.Location = New Point(414, 3)
        BtnClose.Name = "BtnClose"
        BtnClose.Size = New Size(44, 28)
        BtnClose.TabIndex = 0
        BtnClose.TooltipText = ""
        BtnClose.UseVisualStyleBackColor = True
        ' 
        ' TlpBottomBar
        ' 
        TlpBottomBar.BackColor = SystemColors.Control
        TlpBottomBar.ColumnCount = 5
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100F))
        TlpBottomBar.Dock = DockStyle.Fill
        TlpBottomBar.Location = New Point(3, 211)
        TlpBottomBar.Name = "TlpBottomBar"
        TlpBottomBar.RowCount = 1
        TlpBottomBar.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TlpBottomBar.Size = New Size(461, 44)
        TlpBottomBar.TabIndex = 1
        ' 
        ' TlpBody
        ' 
        TlpBody.ColumnCount = 2
        TlpBody.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 125F))
        TlpBody.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TlpBody.Controls.Add(PbxIcon, 0, 1)
        TlpBody.Controls.Add(LblTitle, 0, 0)
        TlpBody.Controls.Add(PnMessage, 1, 1)
        TlpBody.Dock = DockStyle.Fill
        TlpBody.Location = New Point(3, 43)
        TlpBody.Name = "TlpBody"
        TlpBody.RowCount = 2
        TlpBody.RowStyles.Add(New RowStyle(SizeType.Absolute, 30F))
        TlpBody.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TlpBody.Size = New Size(461, 162)
        TlpBody.TabIndex = 2
        ' 
        ' PbxIcon
        ' 
        PbxIcon.Dock = DockStyle.Fill
        PbxIcon.Location = New Point(3, 33)
        PbxIcon.Name = "PbxIcon"
        PbxIcon.Size = New Size(119, 126)
        PbxIcon.SizeMode = PictureBoxSizeMode.CenterImage
        PbxIcon.TabIndex = 5
        PbxIcon.TabStop = False
        ' 
        ' LblTitle
        ' 
        TlpBody.SetColumnSpan(LblTitle, 2)
        LblTitle.Dock = DockStyle.Fill
        LblTitle.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        LblTitle.Location = New Point(3, 3)
        LblTitle.Margin = New Padding(3)
        LblTitle.Name = "LblTitle"
        LblTitle.Size = New Size(455, 24)
        LblTitle.TabIndex = 4
        LblTitle.Text = "Title"
        LblTitle.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' PnMessage
        ' 
        PnMessage.AutoScroll = True
        PnMessage.Controls.Add(LblMessage)
        PnMessage.Dock = DockStyle.Fill
        PnMessage.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        PnMessage.Location = New Point(128, 33)
        PnMessage.Name = "PnMessage"
        PnMessage.Padding = New Padding(5)
        PnMessage.Size = New Size(330, 126)
        PnMessage.TabIndex = 3
        ' 
        ' LblMessage
        ' 
        LblMessage.AutoSize = True
        LblMessage.BackColor = Color.Transparent
        LblMessage.Location = New Point(5, 5)
        LblMessage.Name = "LblMessage"
        LblMessage.Size = New Size(37, 17)
        LblMessage.TabIndex = 4
        LblMessage.Text = "Body"
        LblMessage.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' CcException
        ' 
        CcException.DropDownBorderColor = SystemColors.HotTrack
        CcException.DropDownControl = Nothing
        CcException.DropDownEnabled = True
        CcException.HostControl = BtnClose
        ' 
        ' FrmMessageBox
        ' 
        AutoScaleMode = AutoScaleMode.None
        BackColor = Color.White
        ClientSize = New Size(467, 258)
        ControlBox = False
        Controls.Add(TlpContainer)
        FormBorderStyle = FormBorderStyle.None
        Name = "FrmMessageBox"
        ShowIcon = False
        Text = "FrmMessageBox"
        TlpContainer.ResumeLayout(False)
        TlpTopBar.ResumeLayout(False)
        TlpBody.ResumeLayout(False)
        CType(PbxIcon, ComponentModel.ISupportInitialize).EndInit()
        PnMessage.ResumeLayout(False)
        PnMessage.PerformLayout()
        ResumeLayout(False)
    End Sub

    Friend WithEvents TlpContainer As TableLayoutPanel
    Friend WithEvents TlpTopBar As TableLayoutPanel
    Friend WithEvents TlpBottomBar As TableLayoutPanel
    Friend WithEvents BtnClose As NoFocusCueButton
    Friend WithEvents TlpBody As TableLayoutPanel
    Friend WithEvents PnMessage As Panel
    Friend WithEvents LblMessage As Label
    Friend WithEvents PbxIcon As PictureBox
    Friend WithEvents LblTitle As Label
    Friend WithEvents CcException As ControlContainer
End Class
