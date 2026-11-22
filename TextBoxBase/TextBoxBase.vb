Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' Provides a TextBox with native placeholder support.
''' </summary>
Public Class TextBoxBase
    Inherits TextBox

    Private Const EM_SETCUEBANNER As Integer = &H1501
    Private _PlaceholderText As String = String.Empty

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> lParam As String) As IntPtr
    End Function

    ''' <summary>
    ''' Gets or sets the text displayed when the control is empty and unfocused.
    ''' </summary>
    <DefaultValue(""), Localizable(True), Category("Appearance"), Description("Text shown when the control is empty and does not have focus.")>
    Public Overridable Property PlaceholderText As String
        Get
            Return _PlaceholderText
        End Get
        Set(Value As String)
            _PlaceholderText = If(Value, String.Empty)
            UpdatePlaceholder()
        End Set
    End Property

    Protected Overrides Sub OnHandleCreated(E As EventArgs)
        MyBase.OnHandleCreated(E)
        UpdatePlaceholder()
    End Sub

    Private Sub UpdatePlaceholder()
        If Not IsHandleCreated OrElse Multiline Then Return
        SendMessage(Handle, EM_SETCUEBANNER, IntPtr.Zero, _PlaceholderText)
    End Sub

End Class