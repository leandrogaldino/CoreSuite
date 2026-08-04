Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time initialization and smart-tag support for <see cref="BusyOverlay"/>.
''' </summary>
Public NotInheritable Class BusyOverlayDesigner
    Inherits ComponentDesigner
    Private _ActionLists As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart-tag actions available for the associated component.
    ''' </summary>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionLists Is Nothing Then _ActionLists = New DesignerActionListCollection From {New BusyOverlayDesignerActionList(Me)}
            Return _ActionLists
        End Get
    End Property
    ''' <summary>
    ''' Assigns the designer root control as the initial target when the component is added to a form.
    ''' </summary>
    ''' <param name="DefaultValues">The default values supplied by the designer.</param>
    Public Overrides Sub InitializeNewComponent(DefaultValues As IDictionary)
        MyBase.InitializeNewComponent(DefaultValues)
        Dim Overlay As BusyOverlay = TryCast(Component, BusyOverlay)
        If Overlay Is Nothing OrElse Overlay.TargetControl IsNot Nothing Then Return
        Dim Host As IDesignerHost = TryCast(GetService(GetType(IDesignerHost)), IDesignerHost)
        If Host Is Nothing Then Return
        Dim RootControl As Control = TryCast(Host.RootComponent, Control)
        If RootControl Is Nothing Then Return
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(Overlay)(NameOf(Overlay.TargetControl))
        If Descriptor IsNot Nothing Then Descriptor.SetValue(Overlay, RootControl)
    End Sub
End Class
