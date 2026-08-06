Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time behavior and smart-tag actions for the <see cref="NavigationView"/> control.
''' </summary>
Public Class NavigationViewControlDesigner
    Inherits ControlDesigner
    Private _ActionLists As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart-tag action lists available for the associated <see cref="NavigationView"/>.
    ''' </summary>
    ''' <returns>A collection containing the NavigationView design-time actions.</returns>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionLists Is Nothing Then _ActionLists = New DesignerActionListCollection From {New NavigationViewControlDesignerActionList(Me)}
            Return _ActionLists
        End Get
    End Property
    ''' <summary>
    ''' Gets the selection rules that allow the control to be moved and resized in the Windows Forms Designer.
    ''' </summary>
    ''' <returns>Move and resize selection rules.</returns>
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Return SelectionRules.Visible Or SelectionRules.Moveable Or SelectionRules.AllSizeable
        End Get
    End Property
End Class
