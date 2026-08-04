Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Exposes the most frequently used <see cref="BusyOverlay"/> properties through its designer smart tag.
''' </summary>
Public NotInheritable Class BusyOverlayDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Overlay As BusyOverlay
    ''' <summary>
    ''' Initializes a new smart-tag action list for the specified designer.
    ''' </summary>
    ''' <param name="Designer">The designer that owns the component.</param>
    Public Sub New(Designer As BusyOverlayDesigner)
        MyBase.New(Designer.Component)
        _Overlay = DirectCast(Designer.Component, BusyOverlay)
    End Sub
    ''' <summary>
    ''' Gets the ordered collection of smart-tag items.
    ''' </summary>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionHeaderItem("Association"),
            New DesignerActionPropertyItem(NameOf(TargetControl), "TargetControl", "Association", "Specifies the control covered by the overlay."),
            New DesignerActionPropertyItem(NameOf(Enabled), "Enabled", "Association", "Enables or disables overlay display."),
            New DesignerActionHeaderItem("Content"),
            New DesignerActionPropertyItem(NameOf(MessageText), "MessageText", "Content", "Specifies the primary message."),
            New DesignerActionPropertyItem(NameOf(DetailText), "DetailText", "Content", "Specifies optional detail text."),
            New DesignerActionPropertyItem(NameOf(IndicatorStyle), "IndicatorStyle", "Content", "Selects the spinner, marquee, determinate progress, or no indicator."),
            New DesignerActionPropertyItem(NameOf(AllowCancellation), "AllowCancellation", "Content", "Displays a cancel button for cancellable operations."),
            New DesignerActionHeaderItem("Appearance"),
            New DesignerActionPropertyItem(NameOf(OverlayColor), "OverlayColor", "Appearance", "Specifies the color applied across the target."),
            New DesignerActionPropertyItem(NameOf(OverlayOpacity), "OverlayOpacity", "Appearance", "Specifies overlay opacity."),
            New DesignerActionPropertyItem(NameOf(ShowContentPanel), "ShowContentPanel", "Appearance", "Draws content on a centered panel."),
            New DesignerActionPropertyItem(NameOf(ContentBackColor), "ContentBackColor", "Appearance", "Specifies the centered panel color."),
            New DesignerActionHeaderItem("Timing"),
            New DesignerActionPropertyItem(NameOf(OperationDisplayDelay), "OperationDisplayDelay", "Timing", "Avoids showing the overlay for very fast RunAsync operations."),
            New DesignerActionPropertyItem(NameOf(MinimumOperationDisplayTime), "MinimumOperationDisplayTime", "Timing", "Avoids a brief overlay flash.")
        }
    End Function
    ''' <summary>
    ''' Gets or sets the control covered by the component.
    ''' </summary>
    Public Property TargetControl As Control
        Get
            Return _Overlay.TargetControl
        End Get
        Set(value As Control)
            SetProperty(NameOf(TargetControl), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the component can display.
    ''' </summary>
    Public Property Enabled As Boolean
        Get
            Return _Overlay.Enabled
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(Enabled), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the primary message.
    ''' </summary>
    Public Property MessageText As String
        Get
            Return _Overlay.MessageText
        End Get
        Set(value As String)
            SetProperty(NameOf(MessageText), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the detail text.
    ''' </summary>
    Public Property DetailText As String
        Get
            Return _Overlay.DetailText
        End Get
        Set(value As String)
            SetProperty(NameOf(DetailText), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the indicator style.
    ''' </summary>
    Public Property IndicatorStyle As BusyOverlayIndicatorStyle
        Get
            Return _Overlay.IndicatorStyle
        End Get
        Set(value As BusyOverlayIndicatorStyle)
            SetProperty(NameOf(IndicatorStyle), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether cancellable operations display a button.
    ''' </summary>
    Public Property AllowCancellation As Boolean
        Get
            Return _Overlay.AllowCancellation
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(AllowCancellation), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the overlay tint color.
    ''' </summary>
    Public Property OverlayColor As Color
        Get
            Return _Overlay.OverlayColor
        End Get
        Set(value As Color)
            SetProperty(NameOf(OverlayColor), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets overlay opacity.
    ''' </summary>
    Public Property OverlayOpacity As Integer
        Get
            Return _Overlay.OverlayOpacity
        End Get
        Set(value As Integer)
            SetProperty(NameOf(OverlayOpacity), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether content uses a centered panel.
    ''' </summary>
    Public Property ShowContentPanel As Boolean
        Get
            Return _Overlay.ShowContentPanel
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ShowContentPanel), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the centered panel color.
    ''' </summary>
    Public Property ContentBackColor As Color
        Get
            Return _Overlay.ContentBackColor
        End Get
        Set(value As Color)
            SetProperty(NameOf(ContentBackColor), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the RunAsync display delay.
    ''' </summary>
    Public Property OperationDisplayDelay As Integer
        Get
            Return _Overlay.OperationDisplayDelay
        End Get
        Set(value As Integer)
            SetProperty(NameOf(OperationDisplayDelay), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum RunAsync display time.
    ''' </summary>
    Public Property MinimumOperationDisplayTime As Integer
        Get
            Return _Overlay.MinimumOperationDisplayTime
        End Get
        Set(value As Integer)
            SetProperty(NameOf(MinimumOperationDisplayTime), value)
        End Set
    End Property
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Overlay)(PropertyName)
        If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        Descriptor.SetValue(_Overlay, Value)
    End Sub
End Class
