Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Design
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Windows.Forms.Design

<DefaultEvent("ColorChanged")>
<DefaultProperty("Color")>
Public Class ColorPicker
    Inherits Control

    Protected Const EXTRASIZE As Integer = 2

    Protected Const TAB_COMMON As String = "common"

    Protected Const TAB_PALETTE As String = "palette"

    Protected Const TAB_SYSTEM As String = "system"

    Protected ColorUI As Control

    Protected Editor As ColorEditor

    Protected Palette As Control

    Protected Tab As TabControl

    Private _AllowTabOut As Boolean = True

    Private _Color As Color = Color.White

    Private _ColorUIWnd As ColorUIWnd

    Private _CustomColors As Color()

    Private _Service As ColorEditorService

    Private _TabOrderPos As Integer = -1

    Private OrigColors As Color()

    Private Siblings As Control()

    Public Sub New()
        MyBase.New()
        ShowEditor()
    End Sub

    Public Event ColorChanged As EventHandler
    Public Shared ReadOnly Property DefaultCustomColors As Color()
        Get
            Dim initialColors As Color() = New Color(15) {}
            For i As Integer = 0 To initialColors.Length - 1
                initialColors(i) = Color.White
            Next
            Return initialColors
        End Get
    End Property

    <DefaultValue(True)>
    Public Property AllowTabOut As Boolean
        Get
            Return _AllowTabOut
        End Get
        Set(ByVal value As Boolean)
            _AllowTabOut = value
            _TabOrderPos = -1
            Siblings = Nothing
        End Set
    End Property

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

    <DefaultValue(GetType(Color), Nothing)>
    Public Property CustomColors As Color()
        Get
            Return _CustomColors
        End Get
        Set(ByVal value As Color())
            ShouldSerializeCustomColors = False
            If value IsNot Nothing Then
                Debug.Assert(value.Length = 16)
                OrigColors = CType(value.Clone(), Color())
            Else
                OrigColors = Nothing
            End If

            _CustomColors = value
            SetEditorCustomColors()
        End Set
    End Property
    Public Overrides Property MinimumSize As Size
        Get
            Return Me.DefaultMinimumSize
        End Get
        Set(ByVal value As Size)
        End Set
    End Property

    <Browsable(False)>
    <DefaultValue(False)>
    Public Property ShouldSerializeCustomColors As Boolean
    Protected Overrides ReadOnly Property DefaultMinimumSize As Size
        Get
            Dim extraSize As Integer = extraSize + 2 * CInt(0)
            Return New Size(202 + extraSize, 220 + extraSize)
        End Get
    End Property

    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return Me.DefaultMinimumSize
        End Get
    End Property

    Private ReadOnly Property TabOrderPos As Integer
        Get
            If _TabOrderPos <> -1 Then
                Return _TabOrderPos
            End If
            _TabOrderPos = -2
            Siblings = Nothing
            Dim Parent As Control = Me.Parent
            If Parent IsNot Nothing AndAlso Parent.Controls.Count > 1 Then
                Siblings = New Control(Parent.Controls.Count - 1) {}
                Parent.Controls.CopyTo(Siblings, 0)
                Dim TabIndices As Integer() = New Integer(Parent.Controls.Count - 1) {}
                For i As Integer = 0 To Siblings.Length - 1
                    TabIndices(i) = Siblings(i).TabIndex
                Next
                Array.Sort(TabIndices, Siblings)
                _TabOrderPos = Array.IndexOf(Siblings, Me)
            End If
            Return _TabOrderPos
        End Get
    End Property

    Public Sub CloseEditor()
        CloseEditorInternal()
        _Service = Nothing
        ColorUI = Nothing
        Palette = Nothing
        Tab = Nothing
        Editor = Nothing
    End Sub

    Public Sub PaintValue(ByVal color As Color, ByVal canvas As Graphics, ByVal rectangle As Rectangle)
        Editor?.PaintValue(color, canvas, rectangle)
    End Sub

    Public Sub ShowEditor()
        ShouldSerializeCustomColors = False
        If _Service Is Nothing Then
            _Service = New ColorEditorService()
            AddHandler _Service.ColorUIAvailable, AddressOf Service_ColorUIAvailable
            AddHandler _Service.ColorChanged, AddressOf Service_ColorChanged
        End If
        If Editor Is Nothing Then
            Editor = New ColorEditor()
        End If
        If ColorUI Is Nothing Then
            Editor.EditValue(_Service, _Color)
            If Not _Color.IsKnownColor Then
                SetPaletteColor(_Color)
            End If
            RestoreEditorServiceReference()
        End If
    End Sub
    <DllImport("user32.dll")>
    Protected Shared Function GetAsyncKeyState(ByVal vKey As Keys) As Short
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Protected Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Protected Overridable Sub AddColorUI(ByVal ColorUI As Control)
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

    Protected Overridable Sub CloseEditorInternal()
        _Service?.CloseDropDownInternal()
        If ColorUI IsNot Nothing Then
            SendKeyDown(Tab.SelectedTab.Controls(0), Keys.[Return])
            Debug.Assert(_Service Is Nothing)
            Debug.Assert(ColorUI Is Nothing)
            Debug.Assert(Editor Is Nothing)
        End If
    End Sub

    Protected Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            CloseEditor()
        End If
        MyBase.Dispose(Disposing)
    End Sub

    Protected Overrides Sub OnClientSizeChanged(ByVal e As EventArgs)
        If ColorUI Is Nothing AndAlso Me.ClientSize <> Me.DefaultMinimumSize Then
            Me.ClientSize = Me.DefaultMinimumSize
        End If
        MyBase.OnClientSizeChanged(e)
    End Sub

    Protected Overridable Sub OnColorChanged(ByVal CewColor As Color)
        If CewColor <> _Color Then
            _Color = CewColor
            RaiseEvent ColorChanged(Me, EventArgs.Empty)
        End If
    End Sub

    Protected Overrides Sub OnFontChanged(ByVal e As EventArgs)
        MyBase.OnFontChanged(e)
        If ColorUI IsNot Nothing Then
            ColorUI.Font = Me.Font
        End If
    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
        MyBase.OnGotFocus(e)
        If ColorUI IsNot Nothing AndAlso Not ColorUI.ContainsFocus Then
            ColorUI.Focus()
        End If
    End Sub

    Protected Sub SendKeyDown(ByVal Control As Control, ByVal Key As Keys)
        Const WM_KEYDOWN As Integer = &H100
        If Control IsNot Nothing Then
            SendMessage(Control.Handle, WM_KEYDOWN, New IntPtr(CInt(Key)), IntPtr.Zero)
        End If
    End Sub

    Protected Sub SetEditorColor(ByVal newColor As Color)
        If ColorUI IsNot Nothing AndAlso newColor <> _Color Then
            _ColorUIWnd.PreventSizing = True
            Editor.EditValue(_Service, newColor)
            RestoreEditorServiceReference()
            _ColorUIWnd.PreventSizing = False
            If Not newColor.IsKnownColor Then
                SetPaletteColor(newColor)
            End If

            ResetControls()
        End If
    End Sub

    Protected Sub SetEditorCustomColors()
        If ColorUI IsNot Nothing AndAlso _CustomColors IsNot Nothing AndAlso _CustomColors.Length = 16 Then
            Dim Type As Type = Palette.[GetType]()
            Dim FieldInfo As FieldInfo = Type.GetField("customColors", BindingFlags.NonPublic Or BindingFlags.Instance)
            FieldInfo.SetValue(Palette, Me._CustomColors)
            Palette.Refresh()
        End If
    End Sub

    Protected Sub SetPaletteColor(ByVal newColor As Color)
        If Palette IsNot Nothing Then
            Dim t As Type = Palette.[GetType]()
            Dim pInfo As PropertyInfo = t.GetProperty("SelectedColor")
            pInfo.SetValue(Palette, newColor, Nothing)
        End If
    End Sub
    Private Sub CompareCustomColors()
        If _CustomColors IsNot Nothing AndAlso OrigColors IsNot Nothing Then
            For i As Integer = 0 To _CustomColors.Length - 1
                If OrigColors(i) <> _CustomColors(i) Then
                    ShouldSerializeCustomColors = True
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Function FocusNextControl(ByVal Forward As Boolean) As Boolean
        Try
            Dim Pos As Integer = TabOrderPos
            If Pos > -1 Then
                If Forward Then
                    If Threading.Interlocked.Increment(Pos) >= Siblings.Length Then
                        Pos = 0
                    End If
                Else
                    If Threading.Interlocked.Decrement(Pos) <= 0 Then
                        Pos = Siblings.Length - 1
                    End If
                End If
                Dim CtrlToSelect As Control = Siblings(Pos)
                CtrlToSelect.Focus()
                Return CtrlToSelect.ContainsFocus
            End If
        Catch
        End Try
        Return False
    End Function

    Private Sub Palette_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        If Not Palette.Focused Then
            Palette.Focus()
        End If
    End Sub

    Private Sub ResetControls()
        Dim PageName As String = Tab.SelectedTab.Name
        Dim ListBox As ListBox
        Select Case PageName
            Case TAB_COMMON, TAB_SYSTEM
                ListBox = CType(Tab.TabPages(If(PageName = TAB_COMMON, TAB_SYSTEM, TAB_COMMON)).Controls(0), ListBox)
                ListBox.SelectedItem = Nothing
                SetPaletteColor(Color.Empty)
            Case TAB_PALETTE
                ListBox = CType(Tab.TabPages(TAB_COMMON).Controls(0), ListBox)
                ListBox.SelectedItem = Nothing
                ListBox = CType(Tab.TabPages(TAB_SYSTEM).Controls(0), ListBox)
                ListBox.SelectedItem = Nothing
        End Select
    End Sub

    Private Sub RestoreEditorServiceReference()
        If ColorUI IsNot Nothing AndAlso _Service IsNot Nothing Then
            Dim Type As Type = ColorUI.[GetType]()
            Dim FieldInfo As FieldInfo = Type.GetField("edSvc", BindingFlags.NonPublic Or BindingFlags.Instance)
            FieldInfo.SetValue(ColorUI, _Service)
        End If
    End Sub

    Private Sub Service_ColorChanged(ByVal sender As Object, ByVal e As EventArgs)
        If ColorUI Is Nothing Then Return
        Dim PageName As String = Tab.SelectedTab.Name
        Dim Value As Object = Nothing
        Select Case PageName
            Case TAB_COMMON, TAB_SYSTEM
                Dim ListBox As ListBox = CType(Tab.SelectedTab.Controls(0), ListBox)
                Value = ListBox.SelectedItem
            Case TAB_PALETTE
                Dim Type As Type = ColorUI.[GetType]()
                Dim PropertyInfo As PropertyInfo = Type.GetProperty("Value")
                Value = PropertyInfo.GetValue(ColorUI, Nothing)
                If Value IsNot Nothing AndAlso CType(Value, Color) <> Color.White Then
                    CompareCustomColors()
                End If
        End Select
        If Value IsNot Nothing Then
            ResetControls()
            OnColorChanged(CType(Value, Color))
        End If
    End Sub
    Private Sub Service_ColorUIAvailable(ByVal sender As Object, ByVal e As ColorPicker.EditorServiceEventArgs)
        If e.ColorUI IsNot Nothing Then
            If ColorUI Is Nothing Then
                AddColorUI(e.ColorUI)
                SetEditorCustomColors()
                _ColorUIWnd = New ColorUIWnd()
                _ColorUIWnd.AssignHandle(ColorUI.Handle)
            End If
        Else
            RemoveHandler _Service.ColorUIAvailable, AddressOf Service_ColorUIAvailable
            RemoveHandler _Service.ColorChanged, AddressOf Service_ColorChanged
            _Service = Nothing
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
    Private Sub Tab_Deselecting(ByVal sender As Object, ByVal e As TabControlCancelEventArgs)
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
    Private Class ColorEditorService
        Implements IServiceProvider, IWindowsFormsEditorService

        Private closeEditor As Boolean

        Public Event ColorChanged As EventHandler

        Public Event ColorUIAvailable As EventHandler(Of EditorServiceEventArgs)
        Public Sub CloseDropDownInternal()
            closeEditor = True
        End Sub

        Private Function IServiceProvider_GetService(serviceType As Type) As Object Implements IServiceProvider.GetService
            If serviceType = GetType(IWindowsFormsEditorService) Then
                Return Me
            End If

            Return Nothing
        End Function

        Private Sub IWindowsFormsEditorService_CloseDropDown() Implements IWindowsFormsEditorService.CloseDropDown
            If Not closeEditor Then
                RaiseEvent ColorChanged(Me, EventArgs.Empty)
            Else
                RaiseEvent ColorUIAvailable(Me, New EditorServiceEventArgs(Nothing))
            End If
        End Sub

        Private Sub IWindowsFormsEditorService_DropDownControl(control As Control) Implements IWindowsFormsEditorService.DropDownControl
            closeEditor = (control Is Nothing)
            RaiseEvent ColorUIAvailable(Me, New EditorServiceEventArgs(control))
        End Sub

        Private Function IWindowsFormsEditorService_ShowDialog(dialog As Form) As DialogResult Implements IWindowsFormsEditorService.ShowDialog
            Throw New Exception("The method or operation is not implemented.")
        End Function

    End Class

    Private Class ColorUIWnd
        Inherits NativeWindow

        Public PreventSizing As Boolean

        Private Const SWP_NOSIZE As Integer = &H1

        Private Const WM_WINDOWPOSCHANGING As Integer = &H46

        Protected Overrides Sub WndProc(ByRef m As Message)
            If PreventSizing AndAlso m.Msg = WM_WINDOWPOSCHANGING Then
                Dim wp As WINDOWPOS = CType(Marshal.PtrToStructure(m.LParam, GetType(WINDOWPOS)), WINDOWPOS)
                wp.flags = wp.flags Or SWP_NOSIZE
                Marshal.StructureToPtr(wp, m.LParam, False)
            End If
            MyBase.WndProc(m)
        End Sub

        <StructLayout(LayoutKind.Sequential)>
        Private Structure WINDOWPOS
            Public hwnd As IntPtr
            Public hwndInsertAfter As IntPtr
            Public x As Integer
            Public y As Integer
            Public cx As Integer
            Public cy As Integer
            Public flags As Integer
        End Structure
    End Class
    Private Class EditorServiceEventArgs
        Inherits EventArgs

        Private ReadOnly _ColorUI As Control

        Public Sub New(ByVal colorUI As Control)
            _ColorUI = colorUI
        End Sub
        Public ReadOnly Property ColorUI As Control
            Get
                Return _ColorUI
            End Get
        End Property

    End Class
End Class