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
        Dim QueriedBoxOrderBy1 As CoreSuite.QueriedBoxOrderBy = New CoreSuite.QueriedBoxOrderBy()
        SearchBox1 = New CoreSuite.SearchBox()
        SuspendLayout()
        ' 
        ' SearchBox1
        ' 
        SearchBox1.Location = New Point(391, 138)
        SearchBox1.Name = "SearchBox1"
        SearchBox1.Query.Limit = Nothing
        SearchBox1.Query.Offset = Nothing
        QueriedBoxOrderBy1.Column.ColumnName = "a"
        QueriedBoxOrderBy1.Direction = CoreSuite.QueryOrderByDirection.Ascending
        SearchBox1.Query.OrderBy.Add(QueriedBoxOrderBy1)
        SearchBox1.Query.Table.Alias = "p"
        SearchBox1.Query.Table.TableName = "person"
        SearchBox1.Size = New Size(100, 23)
        SearchBox1.TabIndex = 0
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(SearchBox1)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents SearchBox1 As CoreSuite.SearchBox

End Class
