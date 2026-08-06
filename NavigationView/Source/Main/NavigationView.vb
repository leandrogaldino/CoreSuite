Imports System.ComponentModel
''' <summary>
''' Displays a configurable navigation pane and manages lazily created <see cref="UserControl"/> pages in a content area.
''' </summary>
''' <remarks>
''' Configure <see cref="Pages"/> in the Windows Forms Designer by assigning each page a <see cref="NavigationPage.ControlType"/>, or register pages in code with a factory when constructor arguments are required.
''' </remarks>
<DefaultEvent("Navigated")>
<DefaultProperty("Pages")>
<Description("Displays navigation buttons and manages lazily created UserControl pages with configurable caching.")>
<Designer(GetType(NavigationViewControlDesigner))>
Partial Public Class NavigationView
    Inherits UserControl
    Private Const CategoryName As String = "NavigationView"
    Private ReadOnly _Pages As NavigationPageCollection
    Private ReadOnly _NavigationPanel As Panel
    Private ReadOnly _NavigationFlow As FlowLayoutPanel
    Private ReadOnly _ContentPanel As Panel
    Private ReadOnly _ToolTip As ToolTip
    Private _SelectedPage As NavigationPage
    Private _AutoNavigateFirstPage As Boolean = True
    Private _NavigationPosition As NavigationPanePosition = NavigationPanePosition.Left
    Private _NavigationWidth As Integer = 220
    Private _ButtonHeight As Integer = 44
    Private _ButtonSpacing As Integer = 2
    Private _NavigationPadding As New Padding(8)
    Private _ButtonPadding As New Padding(12, 0, 12, 0)
    Private _ContentPadding As New Padding(0)
    Private _ImageSize As New Size(20, 20)
    Private _SelectedIndicatorWidth As Integer = 4
    Private _ShowImages As Boolean = True
    Private _ShowToolTips As Boolean = True
    Private _NavigationBackColor As Color = SystemColors.Control
    Private _ContentBackColor As Color = SystemColors.Window
    Private _ButtonBackColor As Color = SystemColors.Control
    Private _ButtonHoverBackColor As Color = SystemColors.ControlLight
    Private _ButtonForeColor As Color = SystemColors.ControlText
    Private _SelectedButtonBackColor As Color = SystemColors.Highlight
    Private _SelectedButtonForeColor As Color = SystemColors.HighlightText
    Private _SelectedIndicatorColor As Color = SystemColors.Highlight
    Private _IsLoaded As Boolean
    Private _IsDisposing As Boolean
    ''' <summary>
    ''' Occurs before a page is displayed and allows the operation to be canceled.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs before a page is displayed and allows navigation to be canceled.")>
    Public Event Navigating As EventHandler(Of NavigationCancelEventArgs)
    ''' <summary>
    ''' Occurs after a page has been displayed successfully.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after a page has been displayed successfully.")>
    Public Event Navigated As EventHandler(Of NavigationEventArgs)
    ''' <summary>
    ''' Occurs after the selected page changes, including when a current page is closed or removed.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after the selected page changes.")>
    Public Event SelectedPageChanged As EventHandler
    ''' <summary>
    ''' Occurs after a page control is created for the first time or recreated.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after a page control is created.")>
    Public Event PageCreated As EventHandler(Of NavigationPageEventArgs)
    ''' <summary>
    ''' Occurs before an explicitly closed page control is disposed and allows the operation to be canceled.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs before an explicitly closed page control is disposed and allows the operation to be canceled.")>
    Public Event PageClosing As EventHandler(Of NavigationPageCancelEventArgs)
    ''' <summary>
    ''' Occurs after a page control is disposed because it was closed, reloaded, removed, or configured for recreation.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after a page control is disposed.")>
    Public Event PageClosed As EventHandler(Of NavigationPageEventArgs)
    ''' <summary>
    ''' Occurs when a page cannot be created or displayed.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when a page cannot be created or displayed.")>
    Public Event NavigationFailed As EventHandler(Of NavigationFailedEventArgs)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationView"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        _Pages = New NavigationPageCollection()
        AddHandler _Pages.Changed, AddressOf PagesChanged
        _ToolTip = New ToolTip()
        _NavigationFlow = New FlowLayoutPanel With {.AutoScroll = True, .BackColor = _NavigationBackColor, .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .Padding = _NavigationPadding, .WrapContents = False}
        AddHandler _NavigationFlow.Layout, AddressOf NavigationFlowLayout
        _NavigationPanel = New Panel With {.BackColor = _NavigationBackColor, .Dock = DockStyle.Left, .Width = _NavigationWidth}
        _NavigationPanel.Controls.Add(_NavigationFlow)
        _ContentPanel = New Panel With {.BackColor = _ContentBackColor, .Dock = DockStyle.Fill, .Padding = _ContentPadding}
        SuspendLayout()
        Controls.Add(_ContentPanel)
        Controls.Add(_NavigationPanel)
        Name = "NavigationView"
        Size = New Size(720, 420)
        MinimumSize = New Size(240, 120)
        ResumeLayout(False)
    End Sub
    ''' <summary>
    ''' Releases every created page, the tooltip, and all internal controls owned by this instance.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing AndAlso Not _IsDisposing Then
            _IsDisposing = True
            RemoveHandler _Pages.Changed, AddressOf PagesChanged
            RemoveHandler _NavigationFlow.Layout, AddressOf NavigationFlowLayout
            For Each Page As NavigationPage In _Pages
                Page.ReleaseControl()
            Next
            _ToolTip.Dispose()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    ''' <summary>
    ''' Performs automatic initial navigation after the control has loaded.
    ''' </summary>
    ''' <param name="E">The event data.</param>
    Protected Overrides Sub OnLoad(E As EventArgs)
        MyBase.OnLoad(E)
        _IsLoaded = True
        If Not IsInDesignMode AndAlso _AutoNavigateFirstPage AndAlso _SelectedPage Is Nothing Then NavigateFirstAvailablePage()
    End Sub
    ''' <summary>
    ''' Updates the internal navigation buttons when the control font changes.
    ''' </summary>
    ''' <param name="E">The event data.</param>
    Protected Overrides Sub OnFontChanged(E As EventArgs)
        MyBase.OnFontChanged(E)
        ApplyButtonAppearance()
    End Sub
    ''' <summary>
    ''' Updates internal page-button availability when the control enabled state changes.
    ''' </summary>
    ''' <param name="E">The event data.</param>
    Protected Overrides Sub OnEnabledChanged(E As EventArgs)
        MyBase.OnEnabledChanged(E)
        ApplyButtonAppearance()
    End Sub
    ''' <summary>
    ''' Applies right-to-left layout changes to the navigation buttons.
    ''' </summary>
    ''' <param name="E">The event data.</param>
    Protected Overrides Sub OnRightToLeftChanged(E As EventArgs)
        MyBase.OnRightToLeftChanged(E)
        ApplyButtonAppearance()
    End Sub
    Private ReadOnly Property IsInDesignMode As Boolean
        Get
            Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse (Site IsNot Nothing AndAlso Site.DesignMode)
        End Get
    End Property
End Class
