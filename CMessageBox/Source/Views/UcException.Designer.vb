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
        TlpContainer = New TableLayoutPanel()
        TlpBottomBar = New TableLayoutPanel()
        LblExceptionTitle = New Label()
        TxtExceptionBody = New TextBox()
        LblNote = New Label()
        TlpContainer.SuspendLayout()
        TlpBottomBar.SuspendLayout()
        SuspendLayout()
        ' 
        ' TlpContainer
        ' 
        TlpContainer.ColumnCount = 1
        TlpContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TlpContainer.Controls.Add(TlpBottomBar, 0, 2)
        TlpContainer.Controls.Add(LblExceptionTitle, 0, 0)
        TlpContainer.Controls.Add(TxtExceptionBody, 0, 1)
        TlpContainer.Dock = DockStyle.Fill
        TlpContainer.Location = New Point(0, 0)
        TlpContainer.Name = "TlpContainer"
        TlpContainer.RowCount = 3
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 40F))
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TlpContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, 50F))
        TlpContainer.Size = New Size(530, 400)
        TlpContainer.TabIndex = 0
        ' 
        ' TlpBottomBar
        ' 
        TlpBottomBar.BackColor = SystemColors.Control
        TlpBottomBar.ColumnCount = 1
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 20F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 20F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 20F))
        TlpBottomBar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 20F))
        TlpBottomBar.Controls.Add(LblNote, 0, 0)
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
        ' LblNote
        ' 
        LblNote.Dock = DockStyle.Fill
        LblNote.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LblNote.Location = New Point(3, 3)
        LblNote.Margin = New Padding(3)
        LblNote.Name = "LblNote"
        LblNote.Size = New Size(518, 38)
        LblNote.TabIndex = 1
        LblNote.Text = "Um e-mail foi encaminhado ao suporte técnico contendo todos os detalhes relacionados ao erro ocorrido, incluindo as informações necessárias para análise e diagnóstico do problema."
        LblNote.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' UcException
        ' 
        AutoScaleMode = AutoScaleMode.None
        Controls.Add(TlpContainer)
        Name = "UcException"
        Size = New Size(530, 400)
        TlpContainer.ResumeLayout(False)
        TlpContainer.PerformLayout()
        TlpBottomBar.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

    Friend WithEvents TlpContainer As TableLayoutPanel
    Friend WithEvents LblExceptionTitle As Label
    Friend WithEvents TxtExceptionBody As TextBox
    Friend WithEvents TlpBottomBar As TableLayoutPanel
    Friend WithEvents LblNote As Label

End Class
