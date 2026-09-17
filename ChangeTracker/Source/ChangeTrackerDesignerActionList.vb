Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>Displays tracking configuration guidance in the designer smart tag.</summary>
Public Class ChangeTrackerDesignerActionList
    Inherits DesignerActionList
    ''' <summary>Initializes the action list for a tracker designer.</summary>
    ''' <param name="Designer">The designer owning this action list.</param>
    Public Sub New(Designer As ChangeTrackerDesigner)
        MyBase.New(Designer.Component)
    End Sub
    ''' <summary>Gets the ordered smart-tag guidance items.</summary>
    ''' <returns>The steps required to configure and initialize tracking.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionHeaderItem("Control configuration"),
            New DesignerActionTextItem("Select a control and set TrackChanges to True.", "Control configuration"),
            New DesignerActionTextItem("Set TrackedProperties to one or more property names.", "Control configuration"),
            New DesignerActionTextItem("Use TrackedEvents for alternate EventHandler notifications.", "Control configuration"),
            New DesignerActionHeaderItem("Runtime initialization"),
            New DesignerActionTextItem("Call AcceptChanges after loading the initial values.", "Runtime initialization"),
            New DesignerActionTextItem("Handle HasChangesChanged to update the save button.", "Runtime initialization")
        }
    End Function
End Class
