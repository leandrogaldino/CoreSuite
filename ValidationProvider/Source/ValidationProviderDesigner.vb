Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides design-time behavior and smart-tag actions for the <see cref="ValidationProvider"/> component.
''' </summary>
Public Class ValidationProviderDesigner
    Inherits ComponentDesigner
    Private _ActionLists As DesignerActionListCollection
    ''' <summary>
    ''' Gets the smart-tag action lists available for the associated provider.
    ''' </summary>
    ''' <returns>A collection containing the provider design-time actions.</returns>
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            If _ActionLists Is Nothing Then _ActionLists = New DesignerActionListCollection From {New ValidationProviderDesignerActionList(Me)}
            Return _ActionLists
        End Get
    End Property
    ''' <summary>
    ''' Initializes a newly created provider and associates it with the designer root container.
    ''' </summary>
    ''' <param name="DefaultValues">The default property values supplied by the designer.</param>
    Public Overrides Sub InitializeNewComponent(DefaultValues As IDictionary)
        MyBase.InitializeNewComponent(DefaultValues)
        Dim Provider As ValidationProvider = TryCast(Component, ValidationProvider)
        If Provider Is Nothing OrElse Provider.ContainerControl IsNot Nothing Then Return
        Dim Host As IDesignerHost = TryCast(GetService(GetType(IDesignerHost)), IDesignerHost)
        If Host Is Nothing Then Return
        Dim RootContainer As ContainerControl = TryCast(Host.RootComponent, ContainerControl)
        If RootContainer Is Nothing Then Return
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(Provider)(NameOf(Provider.ContainerControl))
        If Descriptor IsNot Nothing Then Descriptor.SetValue(Provider, RootContainer)
    End Sub
End Class
