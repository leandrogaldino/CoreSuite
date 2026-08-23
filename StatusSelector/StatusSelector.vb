Imports System.ComponentModel
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
''' <summary>
''' Provides a reusable ToolStrip dropdown selector for application statuses and states.
''' </summary>
''' <remarks>
''' Each item associates display text and appearance with an arbitrary application value. The selected item is reflected by the button and can be retrieved through <see cref="SelectedValue"/>.
''' </remarks>
<DefaultEvent("SelectedValueChanged"), DefaultProperty("SelectedValue")>
Public Class StatusSelector
    Inherits ToolStripDropDownButton
    Private ReadOnly _Items As StatusSelectorItemCollection
    Private _SelectedItem As StatusSelectorItem
    Private _MenuBackColor As Color = Color.White
    Private _MenuBorderColor As Color = Color.FromArgb(210, 210, 210)
    Private _HoverBackColor As Color = Color.FromArgb(240, 247, 255)
    Private _HoverForeColor As Color = Color.Empty
    Private _HoverBorderColor As Color = Color.FromArgb(210, 230, 250)
    Private _ImageMarginBackColor As Color = Color.White
    Private _SelectionIndicatorColor As Color = Color.DodgerBlue
    Private _SelectionIndicatorThickness As Single = 2.0F
    Private _ShowSelectedCheckMark As Boolean = True
    Private _UseSelectedItemForeColor As Boolean = True
    Private _UseSelectedItemBackColor As Boolean
    Private _UnselectedText As String = "Selecione..."
    Private _UnselectedForeColor As Color = SystemColors.ControlText
    Private _UnselectedBackColor As Color = Color.Empty
    Private _MenuItemPadding As New Padding(6, 4, 6, 4)
    Private _MenuPadding As New Padding(2)
    Private _MenuMinimumWidth As Integer
    Private _AutoSizeDropDownWidth As Boolean = True
    Private _MenuRoundedEdges As Boolean
    Private _ItemsUpdateCount As Integer
    ''' <summary>
    ''' Occurs when the selected item changes.
    ''' </summary>
    <Category("StatusSelector"), Description("Occurs when the selected status item changes.")>
    Public Event SelectedItemChanged As EventHandler
    ''' <summary>
    ''' Occurs when the selected item index changes.
    ''' </summary>
    <Category("StatusSelector"), Description("Occurs when the selected status item index changes.")>
    Public Event SelectedIndexChanged As EventHandler
    ''' <summary>
    ''' Occurs when the selected value changes.
    ''' </summary>
    <Category("StatusSelector"), Description("Occurs when the selected status value changes.")>
    Public Event SelectedValueChanged As EventHandler
    ''' <summary>
    ''' Occurs when a status item is clicked.
    ''' </summary>
    <Category("StatusSelector"), Description("Occurs when a status item is clicked.")>
    Public Event StatusItemClick As EventHandler(Of StatusSelectorItemEventArgs)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="StatusSelector"/> class.
    ''' </summary>
    ''' <remarks>
    ''' The selector uses text-only display by default. Images assigned to status items remain available and can be displayed by changing <see cref="DisplayStyle"/> to <see cref="ToolStripItemDisplayStyle.ImageAndText"/>.
    ''' </remarks>
    Public Sub New()
        _Items = New StatusSelectorItemCollection(Me)
        MyBase.DisplayStyle = ToolStripItemDisplayStyle.Text
        MyBase.Image = Nothing
        Text = _UnselectedText
        ForeColor = _UnselectedForeColor
        DropDown.BackColor = _MenuBackColor
        ApplyMenuAppearance()
    End Sub
    ''' <summary>
    ''' Gets or sets how text and images are displayed by the selector button.
    ''' </summary>
    ''' <remarks>
    ''' The default value is <see cref="ToolStripItemDisplayStyle.Text"/>. Use <see cref="ToolStripItemDisplayStyle.ImageAndText"/> to display the selected status item image beside its text.
    ''' </remarks>
    <Category("Appearance"), Description("Specifies whether the selector displays text, images, or both."), DefaultValue(ToolStripItemDisplayStyle.Text)>
    Public Shadows Property DisplayStyle As ToolStripItemDisplayStyle
        Get
            Return MyBase.DisplayStyle
        End Get
        Set(value As ToolStripItemDisplayStyle)
            If MyBase.DisplayStyle = value Then Return
            MyBase.DisplayStyle = value
        End Set
    End Property
    ''' <summary>
    ''' Gets the collection of status items displayed by the selector.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Items As StatusSelectorItemCollection
        Get
            Return _Items
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the currently selected item.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedItem As StatusSelectorItem
        Get
            Return _SelectedItem
        End Get
        Set(value As StatusSelectorItem)
            If value IsNot Nothing AndAlso Not _Items.Contains(value) Then Throw New ArgumentException("The selected item must belong to this StatusSelector.", NameOf(value))
            SetSelectedItem(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the index of the currently selected item. A value of -1 clears the selection.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedIndex As Integer
        Get
            Return If(_SelectedItem Is Nothing, -1, _Items.IndexOf(_SelectedItem))
        End Get
        Set(value As Integer)
            If value < -1 OrElse value >= _Items.Count Then Throw New ArgumentOutOfRangeException(NameOf(value))
            SelectedItem = If(value = -1, Nothing, _Items(value))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the application value represented by the selected item.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedValue As Object
        Get
            Return _SelectedItem?.Value
        End Get
        Set(value As Object)
            If value Is Nothing Then
                ClearSelection()
                Return
            End If
            Dim StatusItem = FindByValue(value)
            If StatusItem Is Nothing Then
                ClearSelection()
                Return
            End If
            SelectedItem = StatusItem
        End Set
    End Property
    ''' <summary>
    ''' Gets the text of the currently selected item.
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property SelectedText As String
        Get
            Return If(_SelectedItem?.Text, String.Empty)
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the dropdown menu background color.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the dropdown menu background color.")>
    Public Property MenuBackColor As Color
        Get
            Return _MenuBackColor
        End Get
        Set(value As Color)
            If _MenuBackColor = value Then Return
            _MenuBackColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the dropdown menu border color.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the dropdown menu border color.")>
    Public Property MenuBorderColor As Color
        Get
            Return _MenuBorderColor
        End Get
        Set(value As Color)
            If _MenuBorderColor = value Then Return
            _MenuBorderColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default item background color while the pointer is over an item.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the default hover background color.")>
    Public Property HoverBackColor As Color
        Get
            Return _HoverBackColor
        End Get
        Set(value As Color)
            If _HoverBackColor = value Then Return
            _HoverBackColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default item foreground color while the pointer is over an item.
    ''' </summary>
    ''' <remarks>
    ''' <see cref="Color.Empty"/> preserves the foreground color assigned to each individual status item.
    ''' </remarks>
    <Category("StatusSelector Appearance"), Description("Specifies the default hover foreground color. Color.Empty preserves each item foreground color.")>
    Public Property HoverForeColor As Color
        Get
            Return _HoverForeColor
        End Get
        Set(value As Color)
            If _HoverForeColor = value Then Return
            _HoverForeColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the border color drawn around the item currently under the pointer.
    ''' </summary>
    ''' <remarks>
    ''' <see cref="Color.Empty"/> disables the hover border.
    ''' </remarks>
    <Category("StatusSelector Appearance"), Description("Specifies the hover border color. Color.Empty disables the hover border.")>
    Public Property HoverBorderColor As Color
        Get
            Return _HoverBorderColor
        End Get
        Set(value As Color)
            If _HoverBorderColor = value Then Return
            _HoverBorderColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color of the image and selection indicator margin.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the background color of the image and check-mark margin.")>
    Public Property ImageMarginBackColor As Color
        Get
            Return _ImageMarginBackColor
        End Get
        Set(value As Color)
            If _ImageMarginBackColor = value Then Return
            _ImageMarginBackColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color used to draw the selected item check mark.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the selected item check-mark color.")>
    Public Property SelectionIndicatorColor As Color
        Get
            Return _SelectionIndicatorColor
        End Get
        Set(value As Color)
            If _SelectionIndicatorColor = value Then Return
            _SelectionIndicatorColor = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the thickness of the selected item check mark.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the selected item check-mark thickness."), DefaultValue(2.0F)>
    Public Property SelectionIndicatorThickness As Single
        Get
            Return _SelectionIndicatorThickness
        End Get
        Set(value As Single)
            If value <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "Selection indicator thickness must be greater than zero.")
            If Math.Abs(_SelectionIndicatorThickness - value) < Single.Epsilon Then Return
            _SelectionIndicatorThickness = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the selected item is marked with a check mark.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies whether the selected item is marked with a check mark."), DefaultValue(True)>
    Public Property ShowSelectedCheckMark As Boolean
        Get
            Return _ShowSelectedCheckMark
        End Get
        Set(value As Boolean)
            If _ShowSelectedCheckMark = value Then Return
            _ShowSelectedCheckMark = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the button adopts the selected item foreground color.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies whether the selector button uses the selected item foreground color."), DefaultValue(True)>
    Public Property UseSelectedItemForeColor As Boolean
        Get
            Return _UseSelectedItemForeColor
        End Get
        Set(value As Boolean)
            If _UseSelectedItemForeColor = value Then Return
            _UseSelectedItemForeColor = value
            UpdateButtonDisplay()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the button adopts the selected item background color.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies whether the selector button uses the selected item background color."), DefaultValue(False)>
    Public Property UseSelectedItemBackColor As Boolean
        Get
            Return _UseSelectedItemBackColor
        End Get
        Set(value As Boolean)
            If _UseSelectedItemBackColor = value Then Return
            _UseSelectedItemBackColor = value
            UpdateButtonDisplay()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text displayed when no item is selected.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the text displayed when no status is selected."), DefaultValue("Selecione...")>
    Public Property UnselectedText As String
        Get
            Return _UnselectedText
        End Get
        Set(value As String)
            Dim NewValue = If(value, String.Empty)
            If String.Equals(_UnselectedText, NewValue, StringComparison.Ordinal) Then Return
            _UnselectedText = NewValue
            UpdateButtonDisplay()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the button foreground color used when no item is selected.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies the selector foreground color when no status is selected.")>
    Public Property UnselectedForeColor As Color
        Get
            Return _UnselectedForeColor
        End Get
        Set(value As Color)
            If _UnselectedForeColor = value Then Return
            _UnselectedForeColor = value
            UpdateButtonDisplay()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the button background color used when no item is selected.
    ''' </summary>
    ''' <remarks>
    ''' <see cref="Color.Empty"/> leaves the inherited ToolStrip background color unchanged.
    ''' </remarks>
    <Category("StatusSelector Appearance"), Description("Specifies the selector background color when no status is selected. Color.Empty leaves the inherited ToolStrip color unchanged.")>
    Public Property UnselectedBackColor As Color
        Get
            Return _UnselectedBackColor
        End Get
        Set(value As Color)
            If _UnselectedBackColor = value Then Return
            _UnselectedBackColor = value
            UpdateButtonDisplay()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the padding applied to each dropdown item.
    ''' </summary>
    <Category("StatusSelector Layout"), Description("Specifies the padding applied to each dropdown item.")>
    Public Property MenuItemPadding As Padding
        Get
            Return _MenuItemPadding
        End Get
        Set(value As Padding)
            If _MenuItemPadding = value Then Return
            _MenuItemPadding = value
            RebuildDropDown()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the padding applied to the dropdown menu.
    ''' </summary>
    <Category("StatusSelector Layout"), Description("Specifies the padding applied to the dropdown menu.")>
    Public Property MenuPadding As Padding
        Get
            Return _MenuPadding
        End Get
        Set(value As Padding)
            If _MenuPadding = value Then Return
            _MenuPadding = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum width of the dropdown menu in pixels.
    ''' </summary>
    ''' <remarks>
    ''' A value of zero disables the minimum width.
    ''' </remarks>
    <Category("StatusSelector Layout"), Description("Specifies the minimum dropdown width in pixels. Zero disables the minimum width."), DefaultValue(0)>
    Public Property MenuMinimumWidth As Integer
        Get
            Return _MenuMinimumWidth
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "Menu minimum width cannot be negative.")
            If _MenuMinimumWidth = value Then Return
            _MenuMinimumWidth = value
            UpdateDropDownSize()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the dropdown width automatically expands to fit its visible items.
    ''' </summary>
    <Category("StatusSelector Layout"), Description("Specifies whether the dropdown width automatically expands to fit visible items."), DefaultValue(True)>
    Public Property AutoSizeDropDownWidth As Boolean
        Get
            Return _AutoSizeDropDownWidth
        End Get
        Set(value As Boolean)
            If _AutoSizeDropDownWidth = value Then Return
            _AutoSizeDropDownWidth = value
            UpdateDropDownSize()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the professional renderer uses rounded menu edges.
    ''' </summary>
    <Category("StatusSelector Appearance"), Description("Specifies whether the dropdown renderer uses rounded edges."), DefaultValue(False)>
    Public Property MenuRoundedEdges As Boolean
        Get
            Return _MenuRoundedEdges
        End Get
        Set(value As Boolean)
            If _MenuRoundedEdges = value Then Return
            _MenuRoundedEdges = value
            ApplyMenuAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Adds a status item.
    ''' </summary>
    ''' <param name="text">The text displayed for the status.</param>
    ''' <param name="value">The application value associated with the status.</param>
    ''' <returns>The created status item.</returns>
    Public Function Add(text As String, value As Object) As StatusSelectorItem
        Return Items.Add(text, value)
    End Function
    ''' <summary>
    ''' Adds a status item with a custom foreground color.
    ''' </summary>
    ''' <param name="text">The text displayed for the status.</param>
    ''' <param name="value">The application value associated with the status.</param>
    ''' <param name="foreColor">The status foreground color.</param>
    ''' <returns>The created status item.</returns>
    Public Function Add(text As String, value As Object, foreColor As Color) As StatusSelectorItem
        Return Items.Add(text, value, foreColor)
    End Function
    ''' <summary>
    ''' Adds a status item with custom foreground and background colors.
    ''' </summary>
    ''' <param name="text">The text displayed for the status.</param>
    ''' <param name="value">The application value associated with the status.</param>
    ''' <param name="foreColor">The status foreground color.</param>
    ''' <param name="backColor">The status background color.</param>
    ''' <returns>The created status item.</returns>
    Public Function Add(text As String, value As Object, foreColor As Color, backColor As Color) As StatusSelectorItem
        Return Items.Add(text, value, foreColor, backColor)
    End Function
    ''' <summary>
    ''' Clears the current selection without removing any items.
    ''' </summary>
    Public Sub ClearSelection()
        SelectedItem = Nothing
    End Sub
    ''' <summary>
    ''' Finds the first item whose value equals the supplied value.
    ''' </summary>
    ''' <param name="value">The application value to locate.</param>
    ''' <returns>The matching status item, or <see langword="Nothing"/> when no match exists.</returns>
    Public Function FindByValue(value As Object) As StatusSelectorItem
        For Each StatusItem As StatusSelectorItem In _Items
            If Object.Equals(StatusItem.Value, value) Then Return StatusItem
        Next
        Return Nothing
    End Function
    ''' <summary>
    ''' Selects the item whose value equals the supplied value.
    ''' </summary>
    ''' <param name="value">The application value to select.</param>
    ''' <returns><see langword="True"/> when a matching item was found and selected; otherwise, <see langword="False"/>.</returns>
    Public Function SelectValue(value As Object) As Boolean
        Dim StatusItem = FindByValue(value)
        If StatusItem Is Nothing Then Return False
        SelectedItem = StatusItem
        Return True
    End Function
    ''' <summary>
    ''' Attaches a status item to the selector.
    ''' </summary>
    ''' <param name="item">The item to attach.</param>
    Friend Sub AttachItem(item As StatusSelectorItem)
        AddHandler item.PropertyChanged, AddressOf StatusItem_PropertyChanged
    End Sub
    ''' <summary>
    ''' Detaches a status item from the selector.
    ''' </summary>
    ''' <param name="item">The item to detach.</param>
    Friend Sub DetachItem(item As StatusSelectorItem)
        RemoveHandler item.PropertyChanged, AddressOf StatusItem_PropertyChanged
    End Sub
    ''' <summary>
    ''' Suspends dropdown rebuilding while multiple item changes are performed.
    ''' </summary>
    Friend Sub BeginItemsUpdate()
        _ItemsUpdateCount += 1
    End Sub
    ''' <summary>
    ''' Resumes dropdown rebuilding after a grouped item update.
    ''' </summary>
    Friend Sub EndItemsUpdate()
        If _ItemsUpdateCount = 0 Then Return
        _ItemsUpdateCount -= 1
        If _ItemsUpdateCount = 0 Then RebuildDropDown()
    End Sub
    ''' <summary>
    ''' Notifies the selector that its item collection changed.
    ''' </summary>
    Friend Sub ItemsChanged()
        If _SelectedItem IsNot Nothing AndAlso Not _Items.Contains(_SelectedItem) Then SetSelectedItem(Nothing)
        If _ItemsUpdateCount = 0 Then RebuildDropDown()
    End Sub
    ''' <summary>
    ''' Handles the display of the status dropdown menu.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    ''' <remarks>
    ''' Updates the dropdown dimensions immediately before the menu is displayed.
    ''' </remarks>
    Protected Overrides Sub OnDropDownShow(e As EventArgs)
        UpdateDropDownSize()
        MyBase.OnDropDownShow(e)
    End Sub
    ''' <summary>
    ''' Rebuilds the underlying ToolStrip menu items from the status item collection.
    ''' </summary>
    Private Sub RebuildDropDown()
        If IsDisposed Then Return
        DropDownItems.Clear()
        For Each StatusItem As StatusSelectorItem In _Items
            Dim MenuItem As New ToolStripMenuItem(StatusItem.Text) With {
                .Tag = StatusItem,
                .ForeColor = StatusItem.ForeColor,
                .BackColor = ResolveItemBackColor(StatusItem),
                .Enabled = StatusItem.Enabled,
                .Visible = StatusItem.Visible,
                .Image = StatusItem.Image,
                .ToolTipText = StatusItem.ToolTipText,
                .Checked = ReferenceEquals(StatusItem, _SelectedItem),
                .Padding = _MenuItemPadding
            }
            If StatusItem.Font IsNot Nothing Then MenuItem.Font = StatusItem.Font
            AddHandler MenuItem.Click, AddressOf MenuItem_Click
            DropDownItems.Add(MenuItem)
        Next
        ApplyMenuAppearance()
        UpdateButtonDisplay()
    End Sub
    ''' <summary>
    ''' Applies the configured dropdown appearance and renderer.
    ''' </summary>
    Private Sub ApplyMenuAppearance()
        If IsDisposed Then Return
        DropDown.BackColor = _MenuBackColor
        DropDown.Padding = _MenuPadding
        DropDown.Renderer = New StatusSelectorRenderer(Me)
        Dim Menu = TryCast(DropDown, ToolStripDropDownMenu)
        If Menu IsNot Nothing Then
            Menu.ShowCheckMargin = _ShowSelectedCheckMark
            Menu.ShowImageMargin = _Items.Any(Function(StatusItem) StatusItem.Image IsNot Nothing)
        End If
        UpdateDropDownSize()
        DropDown.Invalidate()
    End Sub
    ''' <summary>
    ''' Updates the dropdown width according to its configured layout settings.
    ''' </summary>
    Private Sub UpdateDropDownSize()
        If IsDisposed OrElse DropDownItems.Count = 0 Then Return
        Dim DesiredWidth = _MenuMinimumWidth
        If _AutoSizeDropDownWidth Then
            For Each ToolItem As ToolStripItem In DropDownItems
                If Not ToolItem.Available Then Continue For
                DesiredWidth = Math.Max(DesiredWidth, ToolItem.GetPreferredSize(Size.Empty).Width + DropDown.Padding.Horizontal + 8)
            Next
            DesiredWidth = Math.Max(DesiredWidth, Width)
        End If
        If DesiredWidth > 0 Then
            DropDown.MinimumSize = New Size(DesiredWidth, 0)
        Else
            DropDown.MinimumSize = Size.Empty
        End If
    End Sub
    ''' <summary>
    ''' Changes the currently selected status item and raises the appropriate selection events.
    ''' </summary>
    ''' <param name="value">The new selected item.</param>
    Private Sub SetSelectedItem(value As StatusSelectorItem)
        If ReferenceEquals(_SelectedItem, value) Then Return
        Dim OldValue = _SelectedItem?.Value
        Dim OldIndex = SelectedIndex
        _SelectedItem = value
        UpdateMenuSelection()
        UpdateButtonDisplay()
        RaiseEvent SelectedItemChanged(Me, EventArgs.Empty)
        If OldIndex <> SelectedIndex Then RaiseEvent SelectedIndexChanged(Me, EventArgs.Empty)
        If Not Object.Equals(OldValue, _SelectedItem?.Value) Then RaiseEvent SelectedValueChanged(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Updates the checked state of the dropdown menu items.
    ''' </summary>
    Private Sub UpdateMenuSelection()
        For Each ToolItem As ToolStripItem In DropDownItems
            Dim MenuItem = TryCast(ToolItem, ToolStripMenuItem)
            Dim StatusItem = TryCast(MenuItem?.Tag, StatusSelectorItem)
            If MenuItem Is Nothing OrElse StatusItem Is Nothing Then Continue For
            MenuItem.Checked = ReferenceEquals(StatusItem, _SelectedItem)
        Next
        DropDown.Invalidate()
    End Sub
    ''' <summary>
    ''' Updates the selector button to reflect the current status selection.
    ''' </summary>
    Private Sub UpdateButtonDisplay()
        If _SelectedItem Is Nothing Then
            Text = _UnselectedText
            ForeColor = _UnselectedForeColor
            BackColor = _UnselectedBackColor
            Image = Nothing
            ToolTipText = String.Empty
            Return
        End If
        Text = _SelectedItem.Text
        ForeColor = If(_UseSelectedItemForeColor, _SelectedItem.ForeColor, _UnselectedForeColor)
        BackColor = If(_UseSelectedItemBackColor AndAlso _SelectedItem.BackColor <> Color.Empty, _SelectedItem.BackColor, _UnselectedBackColor)
        Image = _SelectedItem.Image
        ToolTipText = _SelectedItem.ToolTipText
    End Sub
    ''' <summary>
    ''' Resolves the effective background color for a status menu item.
    ''' </summary>
    ''' <param name="item">The status item whose background color is being resolved.</param>
    ''' <returns>The item's custom background color, or the menu background color when none is specified.</returns>
    Private Function ResolveItemBackColor(item As StatusSelectorItem) As Color
        Return If(item.BackColor = Color.Empty, _MenuBackColor, item.BackColor)
    End Function
    ''' <summary>
    ''' Handles clicks on generated dropdown menu items.
    ''' </summary>
    ''' <param name="sender">The menu item that was clicked.</param>
    ''' <param name="e">The event data.</param>
    Private Sub MenuItem_Click(sender As Object, e As EventArgs)
        Dim MenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim StatusItem = TryCast(MenuItem.Tag, StatusSelectorItem)
        If StatusItem Is Nothing Then Return
        SelectedItem = StatusItem
        RaiseEvent StatusItemClick(Me, New StatusSelectorItemEventArgs(StatusItem))
    End Sub
    ''' <summary>
    ''' Handles changes to a status item and rebuilds the dropdown to reflect the new configuration.
    ''' </summary>
    ''' <param name="sender">The status item that changed.</param>
    ''' <param name="e">The property change event data.</param>
    Private Sub StatusItem_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        RebuildDropDown()
    End Sub
End Class