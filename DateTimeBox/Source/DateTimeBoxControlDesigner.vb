Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time behavior for the <see cref="DateTimeBox"/> control.
''' </summary>
Public Class DateTimeBoxControlDesigner
    Inherits ControlDesigner
    Private _ActionList As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart tag action lists available for the associated control.
    ''' </summary>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionList Is Nothing Then _ActionList = New DesignerActionListCollection From {New DateTimeBoxControlDesignerActionList(Me)}
            Return _ActionList
        End Get
    End Property
    ''' <summary>
    ''' Gets the rules that allow the control to be moved and resized horizontally in the designer.
    ''' </summary>
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Return SelectionRules.Visible Or SelectionRules.Moveable Or SelectionRules.LeftSizeable Or SelectionRules.RightSizeable
        End Get
    End Property
    ''' <summary>
    ''' Initializes a newly created control with an empty value.
    ''' </summary>
    ''' <param name="DefaultValues">The default property values supplied by the designer.</param>
    Public Overrides Sub InitializeNewComponent(DefaultValues As IDictionary)
        MyBase.InitializeNewComponent(DefaultValues)
        Dim Box As DateTimeBox = TryCast(Control, DateTimeBox)
        If Box IsNot Nothing Then Box.ClearDateTime()
    End Sub
End Class