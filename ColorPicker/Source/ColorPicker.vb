Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Reflection
''' <summary>
''' Represents a color picker control that hosts the standard Windows Forms
''' <see cref="ColorEditor"/> directly within the control surface instead of
''' displaying it as a drop-down editor.
''' </summary>
''' <remarks>
''' This control provides support for standard, system, and custom colors,
''' exposes the selected color through the <see cref="Color"/> property,
''' and raises the <see cref="ColorChanged"/> event whenever the selected
''' color changes.
''' </remarks>
<DefaultEvent("ColorChanged")>
<DefaultProperty("Color")>
Partial Public Class ColorPicker
    Inherits Control
    Private Service As ColorEditorService
    Private _ColorUIWnd As ColorUIWnd
    Private _Color As Color = Color.White
    Private _CustomColors As Color()
    Private _OrigColors As Color()
    Private _AllowTabOut As Boolean = True
    Private _Siblings As Control()
    Private _TabOrderPos As Integer = -1
    ''' <summary>
    ''' Provides access to the underlying <see cref="ColorEditor"/> instance.
    ''' </summary>
    Protected Editor As ColorEditor
    ''' <summary>
    ''' Represents the hosted color editor user interface.
    ''' </summary>
    Protected ColorUI As Control
    ''' <summary>
    ''' Represents the tab control used by the embedded color editor.
    ''' </summary>
    Protected Tab As TabControl
    ''' <summary>
    ''' Represents the palette control hosted by the embedded color editor.
    ''' </summary>
    Protected Palette As Control
    ''' <summary>
    ''' Identifies the palette tab.
    ''' </summary>
    Protected Const TAB_PALETTE As String = "palette"
    ''' <summary>
    ''' Identifies the common colors tab.
    ''' </summary>
    Protected Const TAB_COMMON As String = "common"
    ''' <summary>
    ''' Identifies the system colors tab.
    ''' </summary>
    Protected Const TAB_SYSTEM As String = "system"
    ''' <summary>
    ''' Defines the extra padding applied to the control size.
    ''' </summary>
    Protected Const EXTRASIZE As Integer = 2
    ''' <summary>
    ''' Occurs when the selected <see cref="Color"/> changes.
    ''' </summary>
    Public Event ColorChanged As EventHandler
    ''' <summary>
    ''' Gets the zero-based position of this control in its parent's tab order.
    ''' </summary>
    ''' <remarks>
    ''' The value is calculated on demand and cached for subsequent accesses.
    ''' </remarks
    Private ReadOnly Property TabOrderPos As Integer
        Get
            If _TabOrderPos <> -1 Then
                Return _TabOrderPos
            End If
            _TabOrderPos = -2
            _Siblings = Nothing
            Dim Parent As Control = Me.Parent
            If Parent IsNot Nothing AndAlso Parent.Controls.Count > 1 Then
                _Siblings = New Control(Parent.Controls.Count - 1) {}
                Parent.Controls.CopyTo(_Siblings, 0)
                Dim TabIndices As Integer() = New Integer(Parent.Controls.Count - 1) {}
                For i As Integer = 0 To _Siblings.Length - 1
                    TabIndices(i) = _Siblings(i).TabIndex
                Next i
                Array.Sort(TabIndices, _Siblings)
                _TabOrderPos = Array.IndexOf(_Siblings, Me)
            End If
            Return _TabOrderPos
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the currently selected color.
    ''' </summary>
    <DefaultValue(GetType(Color), "White")>
    Public Property Color As Color
        Get
            Return _Color
        End Get
        Set(ByVal value As Color)
            SetEditorColor(value)
            OnColorChanged(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the custom colors displayed in the palette.
    ''' </summary>
    ''' <remarks>
    ''' The array must contain exactly 16 colors.
    ''' </remarks>

    <DefaultValue(GetType(Color), Nothing)>
    Public Property CustomColors As Color()
        Get
            Return _CustomColors
        End Get
        Set(ByVal value As Color())
            ShouldSerializeCustomColors = False
            If value IsNot Nothing Then
                Debug.Assert(value.Length = 16)
                _OrigColors = CType(value.Clone(), Color())
            Else
                _OrigColors = Nothing
            End If
            _CustomColors = value
            SetEditorCustomColors()
        End Set
    End Property
    ''' <summary>
    ''' Gets the default set of custom colors used by the control.
    ''' </summary>
    Public Shared ReadOnly Property DefaultCustomColors As Color()
        Get
            Dim InitialColors As Color() = New Color(15) {}
            For i As Integer = 0 To InitialColors.Length - 1
                InitialColors(i) = Color.White
            Next i
            Return InitialColors
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the <see cref="CustomColors"/> property should be serialized.
    ''' </summary>
    <Browsable(False)>
    <DefaultValue(False)>
    Public Property ShouldSerializeCustomColors As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether the <c>Tab</c> key can move the focus out of the color editor.
    ''' </summary>
    <DefaultValue(True)>
    Public Property AllowTabOut As Boolean
        Get
            Return _AllowTabOut
        End Get
        Set(ByVal value As Boolean)
            _AllowTabOut = value
            _TabOrderPos = -1
            _Siblings = Nothing
        End Set
    End Property
    ''' <summary>
    ''' Gets the minimum size that can be assigned to the control.
    ''' </summary>
    Public Overrides Property MinimumSize As Size
        Get
            Return Me.DefaultMinimumSize
        End Get
        Set(value As Size)
        End Set
    End Property
    ''' <summary>
    ''' Gets the default size of the control.
    ''' </summary>
    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return Me.DefaultMinimumSize
        End Get
    End Property
    ''' <summary>
    ''' Gets the default minimum size of the control.
    ''' </summary>
    Protected Overrides ReadOnly Property DefaultMinimumSize As Size
        Get
            Dim extraSize As Integer = extraSize + 2 * CInt(0)
            Return New Size(202 + extraSize, 220 + extraSize)
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ColorPicker"/> class.
    ''' </summary>
    Public Sub New()
        MyBase.New()
        ShowEditor()
    End Sub
    ''' <summary>
    ''' Creates and displays the embedded color editor.
    ''' </summary>
    Public Sub ShowEditor()
        ShouldSerializeCustomColors = False
        If Service Is Nothing Then
            Service = New ColorEditorService()
            AddHandler Service.ColorUIAvailable, AddressOf Service_ColorUIAvailable
            AddHandler Service.ColorChanged, AddressOf Service_ColorChanged
        End If
        If Editor Is Nothing Then
            Editor = New ColorEditor()
        End If
        If ColorUI Is Nothing Then
            Editor.EditValue(Service, _Color)
            If Not _Color.IsKnownColor Then
                SetPaletteColor(_Color)
            End If
            RestoreEditorServiceReference()
        End If
    End Sub
    ''' <summary>
    ''' Closes the embedded color editor and releases its resources.
    ''' </summary>
    Public Sub CloseEditor()
        CloseEditorInternal()
        Service = Nothing
        ColorUI = Nothing
        Palette = Nothing
        Tab = Nothing
        Editor = Nothing
    End Sub
    ''' <summary>
    ''' Paints a color preview within the specified rectangle.
    ''' </summary>
    ''' <param name="color">The color to paint.</param>
    ''' <param name="canvas">The graphics surface used for painting.</param>
    ''' <param name="rectangle">The area in which the color preview is drawn.</param>
    Public Sub PaintValue(color As Color, canvas As Graphics, rectangle As Rectangle)
        Editor?.PaintValue(color, canvas, rectangle)
    End Sub
    ''' <summary>
    ''' Updates the color displayed by the embedded editor.
    ''' </summary>
    ''' <param name="newColor">The color to display.</param>
    Protected Sub SetEditorColor(newColor As Color)
        If ColorUI IsNot Nothing AndAlso newColor <> _Color Then
            _ColorUIWnd.PreventSizing = True
            Editor.EditValue(Service, newColor)
            RestoreEditorServiceReference()
            _ColorUIWnd.PreventSizing = False
            If Not newColor.IsKnownColor Then
                SetPaletteColor(newColor)
            End If
            ResetControls()
        End If
    End Sub
    ''' <summary>
    ''' Selects the specified color in the palette tab.
    ''' </summary>
    ''' <param name="newColor">The color to select.</param>
    Protected Sub SetPaletteColor(newColor As Color)
        If Palette IsNot Nothing Then
            Dim Type As Type = Palette.[GetType]()
            Dim Info As PropertyInfo = Type.GetProperty("SelectedColor")
            Info.SetValue(Palette, newColor, Nothing)
        End If
    End Sub
    ''' <summary>
    ''' Applies the current custom colors to the embedded palette.
    ''' </summary>
    Protected Sub SetEditorCustomColors()
        If ColorUI IsNot Nothing AndAlso _CustomColors IsNot Nothing AndAlso _CustomColors.Length = 16 Then
            Dim Type As Type = Palette.[GetType]()
            Dim Info As FieldInfo = Type.GetField("customColors", BindingFlags.NonPublic Or BindingFlags.Instance)
            Info.SetValue(Palette, Me._CustomColors)
            Palette.Refresh()
        End If
    End Sub
    ''' <summary>
    ''' Adds and initializes the embedded color editor user interface.
    ''' </summary>
    ''' <param name="ColorUI">The editor control to host.</param>
    Protected Overridable Sub AddColorUI(ColorUI As Control)
        Me.ColorUI = ColorUI
        Me.Tab = CType(ColorUI.Controls(0), TabControl)
        Me.Palette = Tab.TabPages(0).Controls(0)
        Tab.TabPages(0).Name = TAB_PALETTE
        Tab.TabPages(1).Name = TAB_COMMON
        Tab.TabPages(2).Name = TAB_SYSTEM
        RemoveHandler Tab.Deselecting, AddressOf Tab_Deselecting
        AddHandler Tab.Deselecting, AddressOf Tab_Deselecting
        ColorUI.Font = Me.Font
        ColorUI.Location = New Point(1, 1)
        ColorUI.Size = Me.ClientSize
        RemoveHandler Palette.MouseUp, AddressOf Palette_MouseUp
        AddHandler Palette.MouseUp, AddressOf Palette_MouseUp
        Me.Controls.Add(ColorUI)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ColorChanged"/> event when the selected color changes.
    ''' </summary>
    ''' <param name="newColor">The newly selected color.</param>
    Protected Overridable Sub OnColorChanged(newColor As Color)
        If newColor <> _Color Then
            _Color = newColor
            RaiseEvent ColorChanged(Me, EventArgs.Empty)
        End If
    End Sub
    ''' <summary>
    ''' Closes the embedded editor without disposing the control itself.
    ''' </summary>
    Protected Overridable Sub CloseEditorInternal()
        Service?.CloseDropDownInternal()
        If ColorUI IsNot Nothing AndAlso
   Not ColorUI.IsDisposed AndAlso
   Tab IsNot Nothing AndAlso
   Tab.SelectedTab IsNot Nothing AndAlso
   Tab.SelectedTab.Controls.Count > 0 Then

            SendKeyDown(Tab.SelectedTab.Controls(0), Keys.Return)
        End If
    End Sub
    ''' <summary>
    ''' Sends a <c>WM_KEYDOWN</c> message to the specified control.
    ''' </summary>
    ''' <param name="Control">The target control.</param>
    ''' <param name="key">The key to simulate.</param>
    Protected Shared Sub SendKeyDown(Control As Control, key As Keys)
        Const WM_KEYDOWN As Integer = &H100
        If Control IsNot Nothing Then
            SendMessage(Control.Handle, WM_KEYDOWN, New IntPtr(CInt(key)), IntPtr.Zero)
        End If
    End Sub
    ''' <summary>
    ''' Restores the editor service reference used internally by the embedded color editor.
    ''' </summary>
    Private Sub RestoreEditorServiceReference()
        If ColorUI IsNot Nothing AndAlso Service IsNot Nothing Then
            Dim Type As Type = ColorUI.[GetType]()
            Dim Info As FieldInfo = Type.GetField("edSvc", BindingFlags.NonPublic Or BindingFlags.Instance)
            Info.SetValue(ColorUI, Service)
        End If
    End Sub
    ''' <summary>
    ''' Resets the selection state of the embedded editor controls.
    ''' </summary>
    Private Sub ResetControls()
        Dim PageName As String = Tab.SelectedTab.Name
        Dim LbColor As ListBox
        Select Case PageName
            Case TAB_COMMON, TAB_SYSTEM
                LbColor = CType(Tab.TabPages(If(PageName = TAB_COMMON, TAB_SYSTEM, TAB_COMMON)).Controls(0), ListBox)
                LbColor.SelectedItem = Nothing
                SetPaletteColor(Color.Empty)
            Case TAB_PALETTE
                LbColor = CType(Tab.TabPages(TAB_COMMON).Controls(0), ListBox)
                LbColor.SelectedItem = Nothing
                LbColor = CType(Tab.TabPages(TAB_SYSTEM).Controls(0), ListBox)
                LbColor.SelectedItem = Nothing
        End Select
    End Sub
    ''' <summary>
    ''' Compares the current custom colors with the original values.
    ''' </summary>
    Private Sub CompareCustomColors()
        If _CustomColors IsNot Nothing AndAlso _OrigColors IsNot Nothing Then
            For i As Integer = 0 To _CustomColors.Length - 1
                If _OrigColors(i) <> _CustomColors(i) Then
                    ShouldSerializeCustomColors = True
                    Exit For
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Handles the creation and destruction of the embedded color editor.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Provides access to the editor UI.</param>
    Private Sub Service_ColorUIAvailable(sender As Object, e As EditorServiceEventArgs)
        If e.ColorUI IsNot Nothing Then
            If ColorUI Is Nothing Then
                AddColorUI(e.ColorUI)
                SetEditorCustomColors()
                _ColorUIWnd = New ColorUIWnd()
                _ColorUIWnd.AssignHandle(ColorUI.Handle)
            End If
        Else
            RemoveHandler Service.ColorUIAvailable, AddressOf Service_ColorUIAvailable
            RemoveHandler Service.ColorChanged, AddressOf Service_ColorChanged
            Service = Nothing
            If Me.Controls.Contains(ColorUI) Then
                Me.Controls.Remove(ColorUI)
            End If
            ColorUI = Nothing
            Palette = Nothing
            Tab = Nothing
            Editor = Nothing
            _ColorUIWnd = Nothing
        End If
    End Sub
    ''' <summary>
    ''' Handles changes to the selected color reported by the embedded editor.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Event data.</param>
    Private Sub Service_ColorChanged(sender As Object, e As EventArgs)
        If ColorUI Is Nothing Then Return
        Dim PageName As String = Tab.SelectedTab.Name
        Dim Value As Object = Nothing
        Select Case PageName
            Case TAB_COMMON, TAB_SYSTEM
                Dim lb As ListBox = CType(Tab.SelectedTab.Controls(0), ListBox)
                Value = lb.SelectedItem
            Case TAB_PALETTE
                Dim Type As Type = ColorUI.[GetType]()
                Dim Info As PropertyInfo = Type.GetProperty("Value")
                Value = Info.GetValue(ColorUI, Nothing)
                If Value IsNot Nothing AndAlso CType(Value, Color) <> Color.White Then
                    CompareCustomColors()
                End If
        End Select
        If Value IsNot Nothing Then
            ResetControls()
            OnColorChanged(CType(Value, Color))
        End If
    End Sub
    ''' <summary>
    ''' Handles tab changes to optionally move the focus outside the control.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Provides information about the tab selection.</param>
    Private Sub Tab_Deselecting(sender As Object, e As TabControlCancelEventArgs)
        If _AllowTabOut AndAlso GetAsyncKeyState(Keys.Tab) <> 0 Then
            If (Control.ModifierKeys And (Keys.Alt Or Keys.Control)) = Keys.None Then
                If (Control.ModifierKeys And Keys.Shift) = Keys.None Then
                    If e.TabPageIndex = Tab.TabPages.Count - 1 Then
                        e.Cancel = FocusNextControl(True)
                    End If
                Else
                    If e.TabPageIndex = 0 Then
                        e.Cancel = FocusNextControl(False)
                    End If
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Moves the focus to the next or previous sibling control.
    ''' </summary>
    ''' <param name="Forward">
    ''' <see langword="True"/> to move forward; otherwise, <see langword="False"/>.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> if the target control received focus; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Function FocusNextControl(Forward As Boolean) As Boolean
        Try
            Dim Pos As Integer = TabOrderPos
            If Pos > -1 Then
                If Forward Then
                    If Threading.Interlocked.Increment(Pos) >= _Siblings.Length Then
                        Pos = 0
                    End If
                Else
                    If Threading.Interlocked.Decrement(Pos) <= 0 Then
                        Pos = _Siblings.Length - 1
                    End If
                End If
                Dim CtrlToSelect As Control = _Siblings(Pos)
                CtrlToSelect.Focus()
                Return CtrlToSelect.ContainsFocus
            End If
        Catch
        End Try
        Return False
    End Function
    ''' <summary>
    ''' Ensures the palette receives focus when it is clicked.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub Palette_MouseUp(sender As Object, e As MouseEventArgs)
        If Not Palette.Focused Then
            Palette.Focus()
        End If
    End Sub
    ''' <summary>
    ''' Transfers focus to the embedded editor when the control receives focus.
    ''' </summary>
    ''' <param name="e">Event data.</param>
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e)
        If ColorUI IsNot Nothing AndAlso Not ColorUI.ContainsFocus Then
            ColorUI.Focus()
        End If
    End Sub
    ''' <summary>
    ''' Updates the embedded editor font when the control font changes.
    ''' </summary>
    ''' <param name="e">Event data.</param>
    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        If ColorUI IsNot Nothing Then
            ColorUI.Font = Me.Font
        End If
    End Sub
    ''' <summary>
    ''' Maintains the fixed client size required by the embedded editor.
    ''' </summary>
    ''' <param name="e">Event data.</param>
    Protected Overrides Sub OnClientSizeChanged(e As EventArgs)
        If ColorUI Is Nothing AndAlso Me.ClientSize <> Me.DefaultMinimumSize Then
            Me.ClientSize = Me.DefaultMinimumSize
        End If
        MyBase.OnClientSizeChanged(e)
    End Sub
    ''' <summary>
    ''' Releases the resources used by the control.
    ''' </summary>
    ''' <param name="disposing">
    ''' <see langword="True"/> to release managed and unmanaged resources; <see langword="False"/> to release only unmanaged resources.
    ''' </param>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            CloseEditor()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class