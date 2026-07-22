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
        ControlContainer1 = New CoreSuite.ControlContainer()
        ControlSlider1 = New CoreSuite.ControlSlider()
        SuspendLayout()
        ' 
        ' ControlContainer1
        ' 
        ControlContainer1.DropDownBorderColor = SystemColors.HotTrack
        ControlContainer1.DropDownControl = Nothing
        ControlContainer1.DropDownEnabled = True
        ControlContainer1.HostControl = Nothing
        ' 
        ' ControlSlider1
        ' 
        ControlSlider1.Child = Me
        ControlSlider1.Parent = Me
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)
    End Sub

    Friend WithEvents ControlContainer1 As CoreSuite.ControlContainer
    Friend WithEvents ControlSlider1 As CoreSuite.ControlSlider

End Class
