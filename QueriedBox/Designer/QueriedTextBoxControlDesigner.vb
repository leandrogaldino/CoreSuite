
Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
Public Class QueriedTextBoxControlDesigner
    Inherits ControlDesigner
    Private _ActionList As DesignerActionListCollection
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionList Is Nothing Then
                _ActionList = New DesignerActionListCollection From {
                    New QueriedTextBoxControlDesignerActionList(Me)
                }
            End If
            Return _ActionList
        End Get
    End Property
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Dim Rules As SelectionRules = SelectionRules.Visible Or SelectionRules.Moveable
            Dim Box = TryCast(Control, QueriedBox)
            If Box IsNot Nothing AndAlso Box.Multiline Then
                Rules = Rules Or SelectionRules.AllSizeable
            Else
                Rules = Rules Or SelectionRules.LeftSizeable Or SelectionRules.RightSizeable
            End If
            Return Rules
        End Get
    End Property
    Public Overrides Sub InitializeNewComponent(defaultValues As System.Collections.IDictionary)
        MyBase.InitializeNewComponent(defaultValues)
        If Control IsNot Nothing Then
            Control.Text = String.Empty
        End If
    End Sub
End Class