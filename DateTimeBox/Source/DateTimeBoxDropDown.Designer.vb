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
        Calendar = New MonthCalendar()
        TimePicker = New DateTimePicker()
        ConfirmButton = New Button()
        CancelButton = New Button()
        TimeLabel = New Label()
        ButtonPanel = New FlowLayoutPanel()
        TimePanel = New TableLayoutPanel()
        ButtonPanel.SuspendLayout()
        TimePanel.SuspendLayout()
        SuspendLayout()
        '
        ' Calendar
        '
        Calendar.Location = New Point(6, 6)
        Calendar.MaxSelectionCount = 1
        Calendar.Name = "Calendar"
        Calendar.ShowToday = True
        Calendar.ShowTodayCircle = True
        Calendar.TabIndex = 0
        '
        ' TimePicker
        '
        TimePicker.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        TimePicker.CustomFormat = "HH:mm"
        TimePicker.Format = DateTimePickerFormat.Custom
        TimePicker.Location = New Point(48, 3)
        TimePicker.Name = "TimePicker"
        TimePicker.ShowUpDown = True
        TimePicker.Size = New Size(175, 23)
        TimePicker.TabIndex = 1
        '
        ' ConfirmButton
        '
        ConfirmButton.AutoSize = True
        ConfirmButton.Location = New Point(65, 3)
        ConfirmButton.Margin = New Padding(3)
        ConfirmButton.MinimumSize = New Size(75, 27)
        ConfirmButton.Name = "ConfirmButton"
        ConfirmButton.Size = New Size(75, 27)
        ConfirmButton.TabIndex = 2
        ConfirmButton.Text = "OK"
        ConfirmButton.UseVisualStyleBackColor = True
        '
        ' CancelButton
        '
        CancelButton.AutoSize = True
        CancelButton.DialogResult = DialogResult.Cancel
        CancelButton.Location = New Point(146, 3)
        CancelButton.Margin = New Padding(3)
        CancelButton.MinimumSize = New Size(75, 27)
        CancelButton.Name = "CancelButton"
        CancelButton.Size = New Size(75, 27)
        CancelButton.TabIndex = 3
        CancelButton.Text = "Cancel"
        CancelButton.UseVisualStyleBackColor = True
        '
        ' TimeLabel
        '
        TimeLabel.Anchor = AnchorStyles.Left
        TimeLabel.AutoSize = True
        TimeLabel.Location = New Point(3, 7)
        TimeLabel.Margin = New Padding(3, 0, 6, 0)
        TimeLabel.Name = "TimeLabel"
        TimeLabel.Size = New Size(36, 15)
        TimeLabel.TabIndex = 0
        TimeLabel.Text = "Time:"
        '
        ' ButtonPanel
        '
        ButtonPanel.Controls.Add(ConfirmButton)
        ButtonPanel.Controls.Add(CancelButton)
        ButtonPanel.FlowDirection = FlowDirection.LeftToRight
        ButtonPanel.Location = New Point(6, 205)
        ButtonPanel.Name = "ButtonPanel"
        ButtonPanel.Padding = New Padding(62, 0, 0, 0)
        ButtonPanel.Size = New Size(227, 33)
        ButtonPanel.TabIndex = 2
        ButtonPanel.WrapContents = False
        '
        ' TimePanel
        '
        TimePanel.ColumnCount = 2
        TimePanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        TimePanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        TimePanel.Controls.Add(TimeLabel, 0, 0)
        TimePanel.Controls.Add(TimePicker, 1, 0)
        TimePanel.Location = New Point(6, 170)
        TimePanel.Name = "TimePanel"
        TimePanel.RowCount = 1
        TimePanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        TimePanel.Size = New Size(227, 29)
        TimePanel.TabIndex = 1
        '
        ' DateTimeBoxDropDown
        '
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = SystemColors.Control
        Controls.Add(ButtonPanel)
        Controls.Add(TimePanel)
        Controls.Add(Calendar)
        MinimumSize = New Size(239, 244)
        Name = "DateTimeBoxDropDown"
        Padding = New Padding(6)
        Size = New Size(239, 244)
        ButtonPanel.ResumeLayout(False)
        ButtonPanel.PerformLayout()
        TimePanel.ResumeLayout(False)
        TimePanel.PerformLayout()
        ResumeLayout(False)
    End Sub
End Class