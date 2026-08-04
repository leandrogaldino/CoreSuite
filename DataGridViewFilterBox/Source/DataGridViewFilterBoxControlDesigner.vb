Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time behavior for the <see cref="DataGridViewFilterBox"/> control.
''' </summary>
Public Class DataGridViewFilterBoxControlDesigner
    Inherits ControlDesigner
    Private _ActionLists As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart tag action lists available for the associated <see cref="DataGridViewFilterBox"/>.
    ''' </summary>
    ''' <returns>A collection containing the filter box design-time actions.</returns>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionLists Is Nothing Then _ActionLists = New DesignerActionListCollection From {New DataGridViewFilterBoxControlDesignerActionList(Me)}
            Return _ActionLists
        End Get
    End Property
    ''' <summary>
    ''' Gets the selection rules that determine how the control can be resized in the Windows Forms designer.
    ''' </summary>
    ''' <returns>Horizontal resizing for a single-line control or full resizing when <see cref="TextBox.Multiline"/> is enabled.</returns>
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Dim Rules As SelectionRules = SelectionRules.Visible Or SelectionRules.Moveable
            Dim FilterBox As DataGridViewFilterBox = TryCast(Control, DataGridViewFilterBox)
            If FilterBox IsNot Nothing AndAlso FilterBox.Multiline Then Return Rules Or SelectionRules.AllSizeable
            Return Rules Or SelectionRules.LeftSizeable Or SelectionRules.RightSizeable
        End Get
    End Property
    ''' <summary>
    ''' Initializes a newly created filter box with empty text.
    ''' </summary>
    ''' <param name="DefaultValues">A dictionary containing the default property values supplied by the designer.</param>
    Public Overrides Sub InitializeNewComponent(DefaultValues As IDictionary)
        MyBase.InitializeNewComponent(DefaultValues)
        If Control IsNot Nothing Then Control.Text = String.Empty
    End Sub
End Class
