Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Reflection
''' <summary>
''' Represents a color picker control that hosts the standard Windows Forms
''' <see cref="ColorEditor"/> directly within the control surface, showing
''' only Web (Common) and System color tabs.
''' </summary>
<DefaultEvent("ColorChanged")>
<DefaultProperty("Color")>
Partial Public Class ColorPicker
    Inherits Control
    Private Service As ColorEditorService
    Private _ColorUIWnd As ColorUIWnd
    Private _Color As Color = Color.White
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
    <Description("Occurs when the selected Color changes")>
    Public Event ColorChanged As EventHandler
    ''' <summary>
    ''' Gets the zero-based position of this control in its parent's tab order.
    ''' </summary>
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
    <Description("Defines the currently selected color displayed by the control.")>
    <Category("ColorPicker")>
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
    ''' Gets or sets a value indicating whether the TAB key can move focus outside the embedded color editor.
    ''' </summary>
    <Description("Determines whether the TAB key can move focus outside the color editor.")>
    <Category("ColorPicker")>
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

    Public Overrides Property MinimumSize As Size
        Get
            Return Me.DefaultMinimumSize
        End Get
        Set(value As Size)
        End Set
    End Property

    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return Me.DefaultMinimumSize
        End Get
    End Property

    Protected Overrides ReadOnly Property DefaultMinimumSize As Size
        Get
            Return New Size(202 + EXTRASIZE, 220 + EXTRASIZE)
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        ShowEditor()
    End Sub

    ''' <summary>
    ''' Creates and displays the embedded color editor.
    ''' </summary>
    Public Sub ShowEditor()
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
        Tab = Nothing
        Editor = Nothing
    End Sub

    ''' <summary>
    ''' Paints a color preview within the specified rectangle.
    ''' </summary>
    Public Sub PaintValue(color As Color, canvas As Graphics, rectangle As Rectangle)
        Editor?.PaintValue(color, canvas, rectangle)
    End Sub

    ''' <summary>
    ''' Updates the color displayed by the embedded editor.
    ''' </summary>
    Protected Sub SetEditorColor(newColor As Color)
        If ColorUI IsNot Nothing AndAlso newColor <> _Color Then
            _ColorUIWnd.PreventSizing = True
            Editor.EditValue(Service, newColor)
            RestoreEditorServiceReference()
            _ColorUIWnd.PreventSizing = False
            ResetControls()
        End If
    End Sub

    ''' <summary>
    ''' Adds and initializes the embedded color editor user interface.
    ''' </summary>
    Protected Overridable Sub AddColorUI(ColorUI As Control)
        Me.ColorUI = ColorUI
        Me.Tab = CType(ColorUI.Controls(0), TabControl)

        ' Mapeia as abas do controle base
        ' Aba 0: Palette | Aba 1: Web (Common) | Aba 2: System
        Dim palettePage As TabPage = Tab.TabPages(0)
        Tab.TabPages(1).Name = TAB_COMMON
        Tab.TabPages(2).Name = TAB_SYSTEM

        ' Remove a aba Personalizado (Palette)
        If Tab.TabPages.Contains(palettePage) Then
            Tab.TabPages.Remove(palettePage)
        End If

        RemoveHandler Tab.Deselecting, AddressOf Tab_Deselecting
        AddHandler Tab.Deselecting, AddressOf Tab_Deselecting

        ColorUI.Font = Me.Font
        ColorUI.Location = New Point(1, 1)
        ColorUI.Size = Me.ClientSize

        Me.Controls.Add(ColorUI)
    End Sub

    Protected Overridable Sub OnColorChanged(newColor As Color)
        If newColor <> _Color Then
            _Color = newColor
            RaiseEvent ColorChanged(Me, EventArgs.Empty)
        End If
    End Sub

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

    Protected Shared Sub SendKeyDown(Control As Control, key As Keys)
        Const WM_KEYDOWN As Integer = &H100
        If Control IsNot Nothing Then
            SendMessage(Control.Handle, WM_KEYDOWN, New IntPtr(CInt(key)), IntPtr.Zero)
        End If
    End Sub

    Private Sub RestoreEditorServiceReference()
        If ColorUI IsNot Nothing AndAlso Service IsNot Nothing Then
            Dim Type As Type = ColorUI.[GetType]()
            Dim Info As FieldInfo = Type.GetField("edSvc", BindingFlags.NonPublic Or BindingFlags.Instance)
            Info?.SetValue(ColorUI, Service)
        End If
    End Sub

    Private Sub ResetControls()
        If Tab Is Nothing OrElse Tab.SelectedTab Is Nothing Then Return

        Dim PageName As String = Tab.SelectedTab.Name
        Dim LbColor As ListBox

        If PageName = TAB_COMMON AndAlso Tab.TabPages.Contains(Tab.TabPages(TAB_SYSTEM)) Then
            LbColor = CType(Tab.TabPages(TAB_SYSTEM).Controls(0), ListBox)
            LbColor.SelectedItem = Nothing
        ElseIf PageName = TAB_SYSTEM AndAlso Tab.TabPages.Contains(Tab.TabPages(TAB_COMMON)) Then
            LbColor = CType(Tab.TabPages(TAB_COMMON).Controls(0), ListBox)
            LbColor.SelectedItem = Nothing
        End If
    End Sub

    Private Sub Service_ColorUIAvailable(sender As Object, e As EditorServiceEventArgs)
        If e.ColorUI IsNot Nothing Then
            If ColorUI Is Nothing Then
                AddColorUI(e.ColorUI)
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
            Tab = Nothing
            Editor = Nothing
            _ColorUIWnd = Nothing
        End If
    End Sub

    Private Sub Service_ColorChanged(sender As Object, e As EventArgs)
        If ColorUI Is Nothing OrElse Tab IsNot Nothing AndAlso Tab.SelectedTab Is Nothing Then Return
        Dim PageName As String = Tab.SelectedTab.Name
            Dim Value As Object = Nothing

            Select Case PageName
                Case TAB_COMMON, TAB_SYSTEM
                    Dim lb As ListBox = CType(Tab.SelectedTab.Controls(0), ListBox)
                    Value = lb.SelectedItem
            End Select

            If Value IsNot Nothing Then
                ResetControls()
                OnColorChanged(CType(Value, Color))
            End If
    End Sub

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

    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e)
        If ColorUI IsNot Nothing AndAlso Not ColorUI.ContainsFocus Then
            ColorUI.Focus()
        End If
    End Sub

    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        If ColorUI IsNot Nothing Then
            ColorUI.Font = Me.Font
        End If
    End Sub

    Protected Overrides Sub OnClientSizeChanged(e As EventArgs)
        If ColorUI Is Nothing AndAlso Me.ClientSize <> Me.DefaultMinimumSize Then
            Me.ClientSize = Me.DefaultMinimumSize
        End If
        MyBase.OnClientSizeChanged(e)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            CloseEditor()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class