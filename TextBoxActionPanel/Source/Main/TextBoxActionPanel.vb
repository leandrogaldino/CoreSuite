Imports System.ComponentModel
''' <summary>
''' Adds a configurable floating action panel to an existing <see cref="TextBoxBase"/> without replacing or subclassing the target control.
''' </summary>
''' <remarks>
''' The component appears in the Windows Forms component tray. Assign <see cref="TargetControl"/>, configure <see cref="Actions"/>, and handle <see cref="ActionClicked"/> or assign an action delegate.
''' </remarks>
<DefaultEvent("ActionClicked")>
<DefaultProperty("Actions")>
<Description("Adds a configurable floating image-action panel to an existing TextBoxBase control.")>
<DesignerCategory("Component")>
<ToolboxItem(True)>
Partial Public Class TextBoxActionPanel
    Inherits Component
    Private Const CategoryName As String = "TextBoxActionPanel"
    Private ReadOnly _Actions As TextBoxActionCollection
    Private ReadOnly _ObservedAncestors As New List(Of Control)
    Private _TargetControl As TextBoxBase
    Private _OwnerForm As Form
    Private _Popup As TextBoxActionPopup
    Private _Enabled As Boolean = True
    Private _ShowOnFocus As Boolean = True
    Private _HideOnLeave As Boolean = True
    Private _Placement As TextBoxActionPanelPlacement = TextBoxActionPanelPlacement.Auto
    Private _ButtonSize As Integer = 24
    Private _ButtonSpacing As Integer
    Private _PanelPadding As Integer
    Private _PanelOffset As Integer
    Private _TransparentBackground As Boolean = True
    Private _ShowBorder As Boolean
    Private _PanelBackColor As Color = SystemColors.Window
    Private _BorderColor As Color = SystemColors.ControlDark
    Private _ButtonBackColor As Color = SystemColors.Window
    Private _ButtonHoverBackColor As Color = SystemColors.ControlLight
    Private _ButtonPressedBackColor As Color = SystemColors.ControlDark
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxActionPanel"/> class.
    ''' </summary>
    Public Sub New()
        _Actions = New TextBoxActionCollection()
        AddHandler _Actions.Changed, AddressOf Actions_Changed
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxActionPanel"/> class and adds it to the specified container.
    ''' </summary>
    ''' <param name="Container">The container that owns the component.</param>
    Public Sub New(Container As IContainer)
        Me.New()
        ArgumentNullException.ThrowIfNull(Container)
        Container.Add(Me)
    End Sub
    ''' <summary>
    ''' Releases the popup window and detaches every event subscribed on the target and its parent hierarchy.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            DetachTargetControl()
            RemoveHandler _Actions.Changed, AddressOf Actions_Changed
            If _Popup IsNot Nothing Then
                RemoveHandler _Popup.ActionClick, AddressOf Popup_ActionClick
                _Popup.Dispose()
                _Popup = Nothing
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Private ReadOnly Property IsInDesignMode As Boolean
        Get
            Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse (Site IsNot Nothing AndAlso Site.DesignMode)
        End Get
    End Property
End Class
