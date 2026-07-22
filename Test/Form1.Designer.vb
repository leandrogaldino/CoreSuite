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
        QueriedBox1 = New CoreSuite.QueriedBox()
        SuspendLayout()
        ' 
        ' QueriedBox1
        ' 
        QueriedBox1.DebugOnTextChanged = False
        QueriedBox1.DisplayFieldAlias = Nothing
        QueriedBox1.DisplayFieldAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.NotSet
        QueriedBox1.DisplayFieldName = Nothing
        QueriedBox1.DisplayMainFieldName = Nothing
        QueriedBox1.DisplayTableAlias = Nothing
        QueriedBox1.DisplayTableName = Nothing
        QueriedBox1.Distinct = False
        QueriedBox1.DropDownAutoStretchRight = False
        QueriedBox1.GridHeaderBackColor = SystemColors.Window
        QueriedBox1.IfNull = Nothing
        QueriedBox1.Location = New Point(273, 121)
        QueriedBox1.MainReturnFieldName = Nothing
        QueriedBox1.MainTableAlias = Nothing
        QueriedBox1.MainTableName = Nothing
        QueriedBox1.Name = "QueriedBox1"
        QueriedBox1.Prefix = Nothing
        QueriedBox1.Size = New Size(100, 23)
        QueriedBox1.Suffix = Nothing
        QueriedBox1.TabIndex = 0
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(QueriedBox1)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents QueriedBox1 As CoreSuite.QueriedBox

End Class
