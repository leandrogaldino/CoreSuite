Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time behavior for the <see cref="AsyncLookupBox"/> control.
''' </summary>
Public Class AsyncLookupBoxControlDesigner
    Inherits ControlDesigner
    Private _ActionLists As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart-tag action lists available for the associated lookup box.
    ''' </summary>
    ''' <returns>A collection containing the lookup-box design-time actions.</returns>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionLists Is Nothing Then _ActionLists = New DesignerActionListCollection From {New AsyncLookupBoxControlDesignerActionList(Me)}
            Return _ActionLists
        End Get
    End Property
    ''' <summary>
    ''' Gets the selection rules that allow horizontal resizing of the single-line lookup control.
    ''' </summary>
    ''' <returns>Move, visibility, and horizontal-resizing rules.</returns>
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Return SelectionRules.Visible Or SelectionRules.Moveable Or SelectionRules.LeftSizeable Or SelectionRules.RightSizeable
        End Get
    End Property
    ''' <summary>
    ''' Initializes a newly created lookup box with empty text.
    ''' </summary>
    ''' <param name="DefaultValues">The default property values supplied by the designer.</param>
    Public Overrides Sub InitializeNewComponent(DefaultValues As IDictionary)
        MyBase.InitializeNewComponent(DefaultValues)
        If Control IsNot Nothing Then Control.Text = String.Empty
    End Sub
End Class
