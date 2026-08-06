Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart-tag property access for the most frequently used <see cref="NavigationView"/> settings.
''' </summary>
Public Class NavigationViewControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As NavigationView
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationViewControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the NavigationView.</param>
    Public Sub New(Designer As NavigationViewControlDesigner)
        MyBase.New(Designer.Component)
        _Control = DirectCast(Designer.Component, NavigationView)
    End Sub
    ''' <summary>
    ''' Gets the page collection exposed by the control.
    ''' </summary>
    Public ReadOnly Property Pages As NavigationPageCollection
        Get
            Return _Control.Pages
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets whether initial navigation occurs automatically.
    ''' </summary>
    Public Property AutoNavigateFirstPage As Boolean
        Get
            Return _Control.AutoNavigateFirstPage
        End Get
        Set(Value As Boolean)
            SetProperty(NameOf(AutoNavigateFirstPage), Value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the navigation pane position.
    ''' </summary>
    Public Property NavigationPosition As NavigationPanePosition
        Get
            Return _Control.NavigationPosition
        End Get
        Set(Value As NavigationPanePosition)
            SetProperty(NameOf(NavigationPosition), Value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the navigation pane width.
    ''' </summary>
    Public Property NavigationWidth As Integer
        Get
            Return _Control.NavigationWidth
        End Get
        Set(Value As Integer)
            SetProperty(NameOf(NavigationWidth), Value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the navigation button height.
    ''' </summary>
    Public Property ButtonHeight As Integer
        Get
            Return _Control.ButtonHeight
        End Get
        Set(Value As Integer)
            SetProperty(NameOf(ButtonHeight), Value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether page images are displayed.
    ''' </summary>
    Public Property ShowImages As Boolean
        Get
            Return _Control.ShowImages
        End Get
        Set(Value As Boolean)
            SetProperty(NameOf(ShowImages), Value)
        End Set
    End Property
    ''' <summary>
    ''' Gets the smart-tag items displayed by the Windows Forms Designer.
    ''' </summary>
    ''' <returns>A collection containing page, behavior, layout, and appearance actions.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionHeaderItem("Pages"),
            New DesignerActionPropertyItem(NameOf(Pages), "Pages", "Pages", "Defines the UserControl pages managed by the NavigationView."),
            New DesignerActionHeaderItem("Behavior"),
            New DesignerActionPropertyItem(NameOf(AutoNavigateFirstPage), "AutoNavigateFirstPage", "Behavior", "Determines whether the first available page opens automatically."),
            New DesignerActionHeaderItem("Layout"),
            New DesignerActionPropertyItem(NameOf(NavigationPosition), "NavigationPosition", "Layout", "Defines the edge that contains the navigation pane."),
            New DesignerActionPropertyItem(NameOf(NavigationWidth), "NavigationWidth", "Layout", "Defines the navigation pane width."),
            New DesignerActionPropertyItem(NameOf(ButtonHeight), "ButtonHeight", "Layout", "Defines the height of each navigation button."),
            New DesignerActionHeaderItem("Appearance"),
            New DesignerActionPropertyItem(NameOf(ShowImages), "ShowImages", "Appearance", "Determines whether page images are displayed.")
        }
    End Function
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Control)(PropertyName)
        If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        Descriptor.SetValue(_Control, Value)
    End Sub
End Class
