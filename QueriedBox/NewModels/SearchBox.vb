Imports System.ComponentModel

Public Class SearchBox
    Inherits TextBox
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Query As New QueriedBoxQuery

End Class
