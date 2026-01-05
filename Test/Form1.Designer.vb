<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.AnimatedBox1 = New CoreSuite.Controls.AnimatedBox()
        Me.SuspendLayout()
        '
        'AnimatedBox1
        '
        Me.AnimatedBox1.Location = New System.Drawing.Point(132, 82)
        Me.AnimatedBox1.Name = "AnimatedBox1"
        Me.AnimatedBox1.ScaleMode = CoreSuite.Controls.AnimationScaleMode.Centrer
        Me.AnimatedBox1.Size = New System.Drawing.Size(310, 234)
        Me.AnimatedBox1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(659, 450)
        Me.Controls.Add(Me.AnimatedBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents AnimatedBox1 As CoreSuite.Controls.AnimatedBox
End Class
