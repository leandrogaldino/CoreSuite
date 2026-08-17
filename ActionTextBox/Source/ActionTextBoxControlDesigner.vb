Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time behavior for the <see cref="ActionTextBox"/> control.
''' </summary>
Public Class ActionTextBoxControlDesigner
    Inherits ControlDesigner
    Private _ActionList As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart tag action lists available for the associated <see cref="ActionTextBox"/> control.
    ''' </summary>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionList Is Nothing Then _ActionList = New DesignerActionListCollection From {New ActionTextBoxControlDesignerActionList(Me)}
            Return _ActionList
        End Get
    End Property
    ''' <summary>
    ''' Gets the selection rules that determine how the associated control can be moved and resized in the Windows Forms designer.
    ''' </summary>
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Dim rules As SelectionRules = SelectionRules.Visible Or SelectionRules.Moveable
            Dim box As ActionTextBox = TryCast(Control, ActionTextBox)
            If box IsNot Nothing AndAlso box.Multiline Then
                rules = rules Or SelectionRules.AllSizeable
            Else
                rules = rules Or SelectionRules.LeftSizeable Or SelectionRules.RightSizeable
            End If
            Return rules
        End Get
    End Property
    ''' <summary>
    ''' Initializes a newly created instance of the associated <see cref="ActionTextBox"/> control.
    ''' </summary>
    Public Overrides Sub InitializeNewComponent(defaultValues As IDictionary)
        MyBase.InitializeNewComponent(defaultValues)
        If Control IsNot Nothing Then Control.Text = String.Empty
    End Sub
End Class
