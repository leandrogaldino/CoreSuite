<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim QueriedBoxField1 As CoreSuite.QueriedBoxField = New CoreSuite.QueriedBoxField()
        QueriedBox1 = New CoreSuite.QueriedBox()
        Button1 = New Button()
        SuspendLayout()
        ' 
        ' QueriedBox1
        ' 
        QueriedBox1.DebugOnTextChanged = False
        QueriedBox1.DisplayFieldAlias = "Nome"
        QueriedBox1.DisplayFieldAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.NotSet
        QueriedBox1.DisplayFieldName = "name"
        QueriedBox1.DisplayMainFieldName = "id"
        QueriedBox1.DisplayTableAlias = ""
        QueriedBox1.DisplayTableName = "person"
        QueriedBox1.Distinct = False
        QueriedBox1.DropDownAutoStretchRight = False
        QueriedBox1.GridHeaderBackColor = SystemColors.Window
        QueriedBox1.IfNull = Nothing
        QueriedBox1.Location = New Point(179, 89)
        QueriedBox1.MainReturnFieldName = "id"
        QueriedBox1.MainTableAlias = ""
        QueriedBox1.MainTableName = "person"
        QueriedBox1.Name = "QueriedBox1"
        QueriedBoxField1.Display = True
        QueriedBoxField1.DisplayFieldAlias = "CPF"
        QueriedBoxField1.DisplayFieldAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.NotSet
        QueriedBoxField1.DisplayFieldName = "document"
        QueriedBoxField1.DisplayMainFieldName = "id"
        QueriedBoxField1.DisplayTableAlias = ""
        QueriedBoxField1.DisplayTableName = "person"
        QueriedBoxField1.Freeze = True
        QueriedBoxField1.IfNull = Nothing
        QueriedBoxField1.Prefix = " - "
        QueriedBoxField1.Suffix = Nothing
        QueriedBox1.OtherFields.Add(QueriedBoxField1)
        QueriedBox1.Prefix = Nothing
        QueriedBox1.Size = New Size(288, 23)
        QueriedBox1.Suffix = Nothing
        QueriedBox1.TabIndex = 0
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(228, 214)
        Button1.Name = "Button1"
        Button1.Size = New Size(75, 23)
        Button1.TabIndex = 1
        Button1.Text = "Button1"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(Button1)
        Controls.Add(QueriedBox1)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents QueriedBox1 As CoreSuite.QueriedBox
    Friend WithEvents Button1 As Button

End Class
