Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time support for the <see cref="QueriedBox"/> control in the Visual Studio designer.
''' </summary>
Public Class QueriedTextBoxControlDesigner
    Inherits ControlDesigner
    Private _ActionList As DesignerActionListCollection
    ''' <summary>
    ''' Gets the collection of smart tag action lists displayed in the designer.
    ''' </summary>
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
    ''' <summary>
    ''' Gets the rules that define how the control can be resized and moved in the designer.
    ''' </summary>
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
    ''' <summary>
    ''' Initializes a newly created component with default design-time values.
    ''' </summary>
    ''' <param name="defaultValues">
    ''' A dictionary containing default values for the component properties.
    ''' </param>
    Public Overrides Sub InitializeNewComponent(defaultValues As IDictionary)
        MyBase.InitializeNewComponent(defaultValues)
        If Control IsNot Nothing Then
            Control.Text = String.Empty
        End If
    End Sub
End Class