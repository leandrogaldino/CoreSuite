Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>Provides smart-tag guidance for the ChangeTracker component.</summary>
Public Class ChangeTrackerDesigner
    Inherits ComponentDesigner
    Private _ActionLists As DesignerActionListCollection
    ''' <summary>Gets the smart-tag action lists associated with the component.</summary>
    ''' <returns>The tracking configuration guidance displayed by the Windows Forms designer.</returns>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionLists Is Nothing Then _ActionLists = New DesignerActionListCollection From {New ChangeTrackerDesignerActionList(Me)}
            Return _ActionLists
        End Get
    End Property
End Class
