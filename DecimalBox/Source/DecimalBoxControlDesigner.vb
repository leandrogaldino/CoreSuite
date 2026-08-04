Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions

''' <summary>
''' Provides design-time behavior for the <see cref="DecimalBox"/> control.
''' </summary>
''' <remarks>
''' This designer supplies smart tag actions, controls the resizing rules
''' available in the Windows Forms designer, and initializes newly created
''' controls with empty text.
''' </remarks>
Public Class DecimalBoxControlDesigner
    Inherits ControlDesigner
    ''' <summary>
    ''' Stores the collection of design-time smart tag action lists.
    ''' </summary>
    Private _ActionList As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart tag action lists available for the associated
    ''' <see cref="DecimalBox"/> control.
    ''' </summary>
    ''' <returns>
    ''' A collection containing the design-time actions provided by
    ''' <see cref="DecimalBoxControlDesignerActionList"/>.
    ''' </returns>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionList Is Nothing Then
                _ActionList = New DesignerActionListCollection From {
                    New DecimalBoxControlDesignerActionList(Me)
                }
            End If
            Return _ActionList
        End Get
    End Property
    ''' <summary>
    ''' Gets the selection rules that determine how the associated control
    ''' can be moved and resized in the Windows Forms designer.
    ''' </summary>
    ''' <returns>
    ''' The selection rules applicable to the associated
    ''' <see cref="DecimalBox"/> control.
    ''' </returns>
    ''' <remarks>
    ''' Multiline controls can be resized in all directions. Single-line
    ''' controls can only be resized horizontally.
    ''' </remarks>
    Public Overrides ReadOnly Property SelectionRules As SelectionRules
        Get
            Dim Rules As SelectionRules = SelectionRules.Visible Or SelectionRules.Moveable
            Dim Box As DecimalBox = TryCast(Control, DecimalBox)
            If Box IsNot Nothing AndAlso Box.Multiline Then
                Rules = Rules Or SelectionRules.AllSizeable
            Else
                Rules = Rules Or SelectionRules.LeftSizeable Or SelectionRules.RightSizeable
            End If
            Return Rules
        End Get
    End Property
    ''' <summary>
    ''' Initializes a newly created instance of the associated
    ''' <see cref="DecimalBox"/> control.
    ''' </summary>
    ''' <param name="DefaultValues">
    ''' A dictionary containing the default property values used to initialize
    ''' the new component.
    ''' </param>
    ''' <remarks>
    ''' The inherited initialization behavior is executed first, after which
    ''' the control text is cleared.
    ''' </remarks>
    Public Overrides Sub InitializeNewComponent(DefaultValues As IDictionary)
        MyBase.InitializeNewComponent(DefaultValues)
        If Control IsNot Nothing Then
            Control.Text = String.Empty
        End If
    End Sub
End Class