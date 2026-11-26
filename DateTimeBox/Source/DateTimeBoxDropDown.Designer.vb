Imports System.Drawing
Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DateTimeBoxDropDown
    Inherits System.Windows.Forms.UserControl
    Private Components As System.ComponentModel.IContainer
    Friend WithEvents Calendar As MonthCalendar
    Friend WithEvents TimePicker As DateTimePicker
    Friend WithEvents ConfirmButton As Button
    Friend WithEvents CancelButton As Button
    Friend WithEvents TimeLabel As Label
    Friend WithEvents ButtonPanel As FlowLayoutPanel
    Friend WithEvents TimePanel As TableLayoutPanel
    ''' <summary>
    ''' Releases the resources used by the dropdown editor.
    ''' </summary>
    ''' <param name="Disposing">
    ''' <see langword="True"/> to release managed and unmanaged resources;
    ''' otherwise, <see langword="False"/>.
    ''' </param>
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        Try
            If Disposing AndAlso Components IsNot Nothing Then Components.Dispose()
        Finally
            MyBase.Dispose(Disposing)
        End Try
    End Sub
    ''' <summary>
    ''' Initializes the controls displayed by the date and time dropdown editor.
    ''' </summary>
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Calendar = New System.Windows.Forms.MonthCalendar()
        Me.TimePicker = New System.Windows.Forms.DateTimePicker()
        Me.ConfirmButton = New System.Windows.Forms.Button()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.TimeLabel = New System.Windows.Forms.Label()
        Me.ButtonPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.TimePanel = New System.Windows.Forms.TableLayoutPanel()
        Me.ButtonPanel.SuspendLayout()
        Me.TimePanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'Calendar
        '
        Me.Calendar.Location = New System.Drawing.Point(5, 5)
        Me.Calendar.Margin = New System.Windows.Forms.Padding(8, 8, 8, 8)
        Me.Calendar.MaxSelectionCount = 1
        Me.Calendar.Name = "Calendar"
        Me.Calendar.TabIndex = 0
        '
        'TimePicker
        '
        Me.TimePicker.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimePicker.CustomFormat = "HH:mm"
        Me.TimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TimePicker.Location = New System.Drawing.Point(44, 3)
        Me.TimePicker.Name = "TimePicker"
        Me.TimePicker.ShowUpDown = True
        Me.TimePicker.Size = New System.Drawing.Size(148, 20)
        Me.TimePicker.TabIndex = 1
        '
        'ConfirmButton
        '
        Me.ConfirmButton.AutoSize = True
        Me.ConfirmButton.Location = New System.Drawing.Point(56, 3)
        Me.ConfirmButton.MinimumSize = New System.Drawing.Size(64, 23)
        Me.ConfirmButton.Name = "ConfirmButton"
        Me.ConfirmButton.Size = New System.Drawing.Size(64, 23)
        Me.ConfirmButton.TabIndex = 2
        Me.ConfirmButton.Text = "OK"
        Me.ConfirmButton.UseVisualStyleBackColor = True
        '
        'CancelButton
        '
        Me.CancelButton.AutoSize = True
        Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton.Location = New System.Drawing.Point(126, 3)
        Me.CancelButton.MinimumSize = New System.Drawing.Size(64, 23)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(64, 23)
        Me.CancelButton.TabIndex = 3
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.TimeLabel.AutoSize = True
        Me.TimeLabel.Location = New System.Drawing.Point(3, 6)
        Me.TimeLabel.Margin = New System.Windows.Forms.Padding(3, 0, 5, 0)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(33, 13)
        Me.TimeLabel.TabIndex = 0
        Me.TimeLabel.Text = "Time:"
        '
        'ButtonPanel
        '
        Me.ButtonPanel.Controls.Add(Me.ConfirmButton)
        Me.ButtonPanel.Controls.Add(Me.CancelButton)
        Me.ButtonPanel.Location = New System.Drawing.Point(5, 178)
        Me.ButtonPanel.Name = "ButtonPanel"
        Me.ButtonPanel.Padding = New System.Windows.Forms.Padding(53, 0, 0, 0)
        Me.ButtonPanel.Size = New System.Drawing.Size(195, 29)
        Me.ButtonPanel.TabIndex = 2
        Me.ButtonPanel.WrapContents = False
        '
        'TimePanel
        '
        Me.TimePanel.ColumnCount = 2
        Me.TimePanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TimePanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TimePanel.Controls.Add(Me.TimeLabel, 0, 0)
        Me.TimePanel.Controls.Add(Me.TimePicker, 1, 0)
        Me.TimePanel.Location = New System.Drawing.Point(5, 147)
        Me.TimePanel.Name = "TimePanel"
        Me.TimePanel.RowCount = 1
        Me.TimePanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TimePanel.Size = New System.Drawing.Size(195, 25)
        Me.TimePanel.TabIndex = 1
        '
        'DateTimeBoxDropDown
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.ButtonPanel)
        Me.Controls.Add(Me.TimePanel)
        Me.Controls.Add(Me.Calendar)
        Me.MinimumSize = New System.Drawing.Size(205, 211)
        Me.Name = "DateTimeBoxDropDown"
        Me.Padding = New System.Windows.Forms.Padding(5, 5, 5, 5)
        Me.Size = New System.Drawing.Size(205, 211)
        Me.ButtonPanel.ResumeLayout(False)
        Me.ButtonPanel.PerformLayout()
        Me.TimePanel.ResumeLayout(False)
        Me.TimePanel.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
End Class
