Imports System.Drawing
Imports System.Windows.Forms
Public Class Separator
    Inherits Control
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        LblSeparator = New Label
        LblSeparator.Text = Nothing
        LblSeparator.AutoSize = False
        LblSeparator.BorderStyle = BorderStyle.Fixed3D
        LblSeparator.Location = New Point(0, 9)
        LblSeparator.Size = New Size(100, 2)
        LblSeparator.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        LblSeparator.BackColor = Color.Red
        Padding = New Padding(0)
        Margin = New Padding(0)
        Size = New Size(100, 18)
        Controls.Add(LblSeparator)
    End Sub
    Friend WithEvents LblSeparator As Label
End Class
