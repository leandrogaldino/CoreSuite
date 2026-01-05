Imports System.Drawing
Imports System.Windows.Forms
Public Class Separator
    Inherits Control
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        LblSeparator = New Label With {
            .Text = Nothing,
            .AutoSize = False,
            .BorderStyle = BorderStyle.Fixed3D,
            .Location = New Point(0, 9),
            .Size = New Size(100, 2),
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .BackColor = Color.Red
        }
        Padding = New Padding(0)
        Margin = New Padding(0)
        Size = New Size(100, 18)
        Controls.Add(LblSeparator)
    End Sub
    Friend WithEvents LblSeparator As Label
End Class
