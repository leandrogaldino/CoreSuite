<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UcException
    Inherits System.Windows.Forms.UserControl

    'O UserControl substitui o descarte para limpar a lista de componentes.
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
        TableLayoutPanel1 = New TableLayoutPanel()
        TlpBottomBar = New TableLayoutPanel()
        LblExceptionTitle = New Label()
        TxtExceptionBody = New TextBox()
        NoFocusCueButton3 = New NoFocusCueButton()
        TableLayoutPanel1.SuspendLayout()
        TlpBottomBar.SuspendLayout()
        SuspendLayout()
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.ColumnCount = 1
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.Controls.Add(TlpBottomBar, 0, 2)
        TableLayoutPanel1.Controls.Add(LblExceptionTitle, 0, 0)
        TableLayoutPanel1.Controls.Add(TxtExceptionBody, 0, 1)
        TableLayoutPanel1.Dock = DockStyle.Fill
        TableLayoutPanel1.Location = New Point(0, 0)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 3
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 40F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 50F))
        TableLayoutPanel1.Size = New Size(530, 400)
        TableLayoutPanel1.TabIndex = 0
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
        TlpBottomBar.Controls.Add(NoFocusCueButton3, 4, 0)
        TlpBottomBar.Dock = DockStyle.Fill
        TlpBottomBar.Location = New Point(3, 353)
        TlpBottomBar.Name = "TlpBottomBar"
        TlpBottomBar.RowCount = 1
        TlpBottomBar.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TlpBottomBar.Size = New Size(524, 44)
        TlpBottomBar.TabIndex = 2
        ' 
        ' LblExceptionTitle
        ' 
        LblExceptionTitle.Dock = DockStyle.Fill
        LblExceptionTitle.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        LblExceptionTitle.Location = New Point(3, 3)
        LblExceptionTitle.Margin = New Padding(3)
        LblExceptionTitle.Name = "LblExceptionTitle"
        LblExceptionTitle.Size = New Size(524, 34)
        LblExceptionTitle.TabIndex = 0
        LblExceptionTitle.Text = "Detalhes do Erro"
        LblExceptionTitle.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' TxtExceptionBody
        ' 
        TxtExceptionBody.BackColor = Color.White
        TxtExceptionBody.Dock = DockStyle.Fill
        TxtExceptionBody.Location = New Point(3, 43)
        TxtExceptionBody.Multiline = True
        TxtExceptionBody.Name = "TxtExceptionBody"
        TxtExceptionBody.ReadOnly = True
        TxtExceptionBody.ScrollBars = ScrollBars.Both
        TxtExceptionBody.Size = New Size(524, 304)
        TxtExceptionBody.TabIndex = 1
        TxtExceptionBody.WordWrap = False
        ' 
        ' NoFocusCueButton3
        ' 
        NoFocusCueButton3.Anchor = AnchorStyles.None
        NoFocusCueButton3.BackColor = Color.White
        NoFocusCueButton3.FlatAppearance.BorderColor = Color.Gainsboro
        NoFocusCueButton3.FlatAppearance.MouseDownBackColor = Color.Silver
        NoFocusCueButton3.FlatAppearance.MouseOverBackColor = Color.LightGray
        NoFocusCueButton3.FlatStyle = FlatStyle.Flat
        NoFocusCueButton3.Location = New Point(427, 6)
        NoFocusCueButton3.Margin = New Padding(3, 6, 3, 3)
        NoFocusCueButton3.Name = "NoFocusCueButton3"
        NoFocusCueButton3.Size = New Size(94, 35)
        NoFocusCueButton3.TabIndex = 2
        NoFocusCueButton3.Text = "Salvar"
        NoFocusCueButton3.TooltipText = ""
        NoFocusCueButton3.UseVisualStyleBackColor = False
        ' 
        ' UcException
        ' 
        AutoScaleMode = AutoScaleMode.None
        Controls.Add(TableLayoutPanel1)
        Name = "UcException"
        Size = New Size(530, 400)
        TableLayoutPanel1.ResumeLayout(False)
        TableLayoutPanel1.PerformLayout()
        TlpBottomBar.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents LblExceptionTitle As Label
    Friend WithEvents TxtExceptionBody As TextBox
    Friend WithEvents TlpBottomBar As TableLayoutPanel
    Friend WithEvents NoFocusCueButton3 As NoFocusCueButton

End Class
