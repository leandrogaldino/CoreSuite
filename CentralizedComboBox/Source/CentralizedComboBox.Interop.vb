Imports System.Runtime.InteropServices
''' <summary>
''' Provides Windows API interop declarations required to customize the internal
''' edit control of a <see cref="ComboBox"/>.
''' </summary>
''' <remarks>
''' A standard Windows ComboBox is composed internally of multiple child controls.
''' When using the editable style, the text input area is a separate native edit
''' control. This partial class contains the required Win32 structures and API
''' declarations to access and modify that internal control.
''' </remarks>
Partial Public Class CentralizedComboBox
    ''' <summary>
    ''' Represents the window style index used by the Win32
    ''' <c>GetWindowLong</c> and <c>SetWindowLong</c> functions.
    ''' </summary>
    Private Const GWL_STYLE As Integer = -16
    ''' <summary>
    ''' Windows edit control style flag that horizontally centers the text.
    ''' </summary>
    Private Const ES_CENTER As Integer = 1
    ''' <summary>
    ''' Represents a Win32 RECT structure containing the coordinates
    ''' of a rectangular area.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        ''' <summary>
        ''' Left coordinate of the rectangle.
        ''' </summary>
        Public Left As Integer
        ''' <summary>
        ''' Top coordinate of the rectangle.
        ''' </summary>
        Public Top As Integer
        ''' <summary>
        ''' Right coordinate of the rectangle.
        ''' </summary>
        Public Right As Integer
        ''' <summary>
        ''' Bottom coordinate of the rectangle.
        ''' </summary>
        Public Bottom As Integer
    End Structure
    ''' <summary>
    ''' Contains information about the internal child windows that compose
    ''' a native ComboBox control.
    ''' </summary>
    ''' <remarks>
    ''' This structure is used by <see cref="GetComboBoxInfo"/> to retrieve
    ''' handles for the ComboBox itself, its editable area and its drop-down list.
    ''' </remarks>
    <StructLayout(LayoutKind.Sequential)>
    Private Structure COMBOBOXINFO
        ''' <summary>
        ''' Size of this structure in bytes.
        ''' </summary>
        Public cbSize As Integer
        ''' <summary>
        ''' Bounding rectangle of the ComboBox item area.
        ''' </summary>
        Public rcItem As RECT
        ''' <summary>
        ''' Bounding rectangle of the drop-down button.
        ''' </summary>
        Public rcButton As RECT
        ''' <summary>
        ''' State information of the drop-down button.
        ''' </summary>
        Public stateButton As Integer
        ''' <summary>
        ''' Window handle of the ComboBox control.
        ''' </summary>
        Public hwndCombo As IntPtr
        ''' <summary>
        ''' Window handle of the internal edit control used by editable ComboBoxes.
        ''' </summary>
        Public hwndEdit As IntPtr
        ''' <summary>
        ''' Window handle of the internal list box used by the drop-down portion.
        ''' </summary>
        Public hwndList As IntPtr
    End Structure
    ''' <summary>
    ''' Retrieves the current style attributes of a specified window.
    ''' </summary>
    ''' <param name="hWnd">
    ''' Handle of the window whose style should be retrieved.
    ''' </param>
    ''' <param name="nIndex">
    ''' Index of the value to retrieve. Uses <see cref="GWL_STYLE"/>.
    ''' </param>
    ''' <returns>
    ''' The current window style value.
    ''' </returns>
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function
    ''' <summary>
    ''' Changes the style attributes of a specified window.
    ''' </summary>
    ''' <param name="hWnd">
    ''' Handle of the window whose style should be modified.
    ''' </param>
    ''' <param name="nIndex">
    ''' Index of the style value to modify. Uses <see cref="GWL_STYLE"/>.
    ''' </param>
    ''' <param name="dwNewLong">
    ''' New style value.
    ''' </param>
    ''' <returns>
    ''' The previous style value before modification.
    ''' </returns>
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function
    ''' <summary>
    ''' Retrieves information about the child windows and visual elements
    ''' contained within a native ComboBox control.
    ''' </summary>
    ''' <param name="hWnd">
    ''' Handle of the ComboBox control.
    ''' </param>
    ''' <param name="pcbi">
    ''' Structure that receives the ComboBox information.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the information was successfully retrieved;
    ''' otherwise, <see langword="False"/>.
    ''' </returns>
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetComboBoxInfo(hWnd As IntPtr, ByRef pcbi As COMBOBOXINFO) As Boolean
    End Function
End Class