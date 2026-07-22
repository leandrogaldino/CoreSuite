Imports System.Runtime.InteropServices
''' <summary>
''' Represents a ComboBox control whose displayed text and selected items
''' are horizontally centered.
''' </summary>
''' <remarks>
''' The standard Windows ComboBox does not expose a property to center the
''' text inside its editable portion. This control combines owner drawing with
''' Win32 API calls to modify the internal edit control style.
''' </remarks>
Partial Public Class CentralizedComboBox
    Inherits ComboBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CentralizedComboBox"/> class.
    ''' </summary>
    Public Sub New()
        MyBase.New()
        DrawMode = DrawMode.OwnerDrawFixed
    End Sub
    ''' <summary>
    ''' Called when the native window handle for the control is created.
    ''' </summary>
    ''' <param name="e">
    ''' Event data associated with the handle creation.
    ''' </param>
    ''' <remarks>
    ''' After the handle exists, the internal native edit control becomes
    ''' available and its style can be modified.
    ''' </remarks>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        SetupInternalEditControl()
    End Sub
    ''' <summary>
    ''' Configures the internal edit control of the ComboBox to center its text.
    ''' </summary>
    ''' <remarks>
    ''' The method retrieves the native edit control handle through
    ''' <see cref="GetComboBoxInfo"/> and applies the
    ''' <c>ES_CENTER</c> Windows style flag.
    ''' </remarks>
    Private Sub SetupInternalEditControl()
        Dim Info As New COMBOBOXINFO()
        Info.cbSize = Marshal.SizeOf(Info)

        If GetComboBoxInfo(Me.Handle, Info) AndAlso Info.hwndEdit <> IntPtr.Zero Then
            Dim Style As Integer = GetWindowLong(Info.hwndEdit, GWL_STYLE)
            Style = Style Or ES_CENTER
            Dim PreviousStyle As Integer = SetWindowLong(Info.hwndEdit, GWL_STYLE, Style)
            If PreviousStyle = 0 Then
                Dim ErrorCode As Integer = Marshal.GetLastWin32Error()
                If ErrorCode <> 0 Then
                    Throw New ComponentModel.Win32Exception(ErrorCode)
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Draws each ComboBox item using custom alignment rules.
    ''' </summary>
    ''' <param name="e">
    ''' Event data containing the drawing surface and item information.
    ''' </param>
    ''' <remarks>
    ''' Items are rendered using <see cref="TextRenderer"/> with both
    ''' horizontal and vertical centering enabled.
    ''' </remarks>
    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        e.DrawBackground()
        If e.Index >= 0 Then
            Dim Txt As String = GetItemText(Items(e.Index))
            Dim Flags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix
            TextRenderer.DrawText(e.Graphics, Txt, e.Font, e.Bounds, e.ForeColor, Flags)
        End If
        e.DrawFocusRectangle()
    End Sub
End Class