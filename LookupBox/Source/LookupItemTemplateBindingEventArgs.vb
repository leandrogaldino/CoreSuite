Imports System.Windows.Forms


''' <summary>Provides the record and reusable UserControl to a template binding handler.</summary>
Public NotInheritable Class LookupItemTemplateBindingEventArgs
    Inherits EventArgs

    Public Sub New(Item As Object, Template As UserControl, ResultIndex As Integer)
        Me.Item = Item
        Me.Template = Template
        Me.ResultIndex = ResultIndex
    End Sub

    Public ReadOnly Property Item As Object
    Public ReadOnly Property Template As UserControl
    Public ReadOnly Property ResultIndex As Integer
End Class
