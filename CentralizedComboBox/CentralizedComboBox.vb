Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' A customized ComboBox that centers text in both the edit control and the dropdown items.
''' </summary>
Public Class CentralizedComboBox
    Inherits ComboBox
    Private Const GWL_STYLE As Integer = -16
    Private Const ES_CENTER As Integer = 1

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure COMBOBOXINFO
        Public cbSize As Integer
        Public rcItem As RECT
        Public rcButton As RECT
        Public stateButton As Integer
        Public hwndCombo As IntPtr
        Public hwndEdit As IntPtr
        Public hwndList As IntPtr
    End Structure
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetComboBoxInfo(ByVal hWnd As IntPtr, ByRef pcbi As COMBOBOXINFO) As Boolean
    End Function

    ''' <summary>
    ''' Initializes a new instance of the <see cref="CentralizedComboBox"/> class.
    ''' Sets DrawMode to OwnerDrawFixed to allow custom item rendering.
    ''' </summary>
    Public Sub New()
        MyBase.New()
        Me.DrawMode = DrawMode.OwnerDrawFixed
    End Sub

    ''' <summary>
    ''' Handles the handle creation process to apply alignment styles to the internal edit control.
    ''' </summary>
    ''' <param name="e">An EventArgs that contains the event data.</param>
    Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
        MyBase.OnHandleCreated(e)
        SetupInternalEditControl()
    End Sub

    ''' <summary>
    ''' Modifies the underlying Win32 Edit control style to center the text.
    ''' </summary>
    Private Sub SetupInternalEditControl()
        Dim Info As New COMBOBOXINFO()
        Info.cbSize = Marshal.SizeOf(Info)
        If GetComboBoxInfo(Me.Handle, Info) AndAlso Info.hwndEdit <> IntPtr.Zero Then
            Dim Style As Integer = GetWindowLong(Info.hwndEdit, GWL_STYLE)
            Style = Style Or ES_CENTER
            SetWindowLong(Info.hwndEdit, GWL_STYLE, Style)
        End If
    End Sub

    ''' <summary>
    ''' Draws the items in the dropdown list with horizontal center alignment.
    ''' </summary>
    ''' <param name="e">A DrawItemEventArgs that contains the event data.</param>
    Protected Overrides Sub OnDrawItem(ByVal e As DrawItemEventArgs)
        e.DrawBackground()
        If e.Index >= 0 Then
            Dim Txt As String = GetItemText(Items(e.Index))
            Dim Flags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or
                                           TextFormatFlags.VerticalCenter Or
                                           TextFormatFlags.NoPrefix
            TextRenderer.DrawText(e.Graphics, Txt, e.Font, e.Bounds, e.ForeColor, Flags)
        End If
        e.DrawFocusRectangle()
    End Sub
End Class