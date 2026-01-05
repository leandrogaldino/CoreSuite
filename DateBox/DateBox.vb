Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
<Designer(GetType(DateBoxDesigner))>
Public Class DateBox
    Inherits MaskedTextBox

    Friend WithEvents ControlContainer As New ControlContainer
    Friend WithEvents Calendar As New MonthCalendar
    Friend WithEvents Button As New PictureBox
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    <EditorBrowsable(EditorBrowsableState.Never), Browsable(False)>
    Overloads ReadOnly Property Mask As String
        Get
            Return MyBase.Mask
        End Get
    End Property

    ''' <summary>
    ''' Define a imagem do botão do calendário.
    ''' </summary>
    <Category("Aparência")>
    <Description("Define a imagem do botão do calendário.")>
    Public Property ButtonImage As Image = My.Resources.Calendar

    ''' <summary>
    ''' Define a data representada pelo controle.
    ''' </summary>
    <Category("Aparência")>
    <Description("Define a data representada pelo controle.")>
    Public Property [Date] As Date?
        Get
            If Text.Replace("/", String.Empty).Trim = String.Empty Then Return Nothing
            If IsDate(Text) Then
                Return CDate(Text)
            Else
                Return Nothing
            End If
        End Get
        Set(value As Date?)
            If value IsNot Nothing Then
                Text = CDate(value).ToString("dd/MM/yyyy")
            Else
                Text = String.Empty
            End If
        End Set
    End Property

    Public Sub New()
        MyBase.Mask = "00/00/0000"
        MinimumSize = New Size(100, 0)
        Button = New PictureBox With {
            .Size = New Size(25, ClientSize.Height + 2)
        }
        Button.Location = New Point(ClientSize.Width - Button.Width + 1, -1)
        Button.Cursor = Cursors.[Default]
        Button.BackgroundImage = My.Resources.Calendar
        Button.TabStop = False
        Button.BackgroundImageLayout = ImageLayout.Center
        Button.BackColor = Color.White
        Controls.Add(Button)
        ControlContainer.HostControl = Button
        Calendar.Visible = False
        Controls.Add(Calendar)
        ControlContainer.DropDownControl = Calendar
    End Sub
    Friend Sub SetMask(Mask As String)
        MyBase.Mask = Mask
    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        Button.Size = New Size(25, ClientSize.Height + 2)
        Button.Location = New Point(ClientSize.Width - Button.Width + 1, -1)
        Button.Cursor = Cursors.[Default]
        Button.BackgroundImage = ButtonImage
        Button.BackgroundImageLayout = ImageLayout.Center
    End Sub
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        Button.BackColor = BackColor

    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)

        If e.KeyCode = Keys.Enter Then

            Calendar.Visible = True
            ControlContainer.ShowDropDown()
            e.SuppressKeyPress = True
        End If
    End Sub
    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        BeginInvoke(CType(Sub()
                              SetMaskedTextBoxSelectAll(CType(Me, MaskedTextBox))
                          End Sub, Action))
    End Sub
    Private Sub SetMaskedTextBoxSelectAll(ByVal TextBox As MaskedTextBox)
        TextBox.SelectAll()
    End Sub
    Protected Overrides Sub OnLeave(e As EventArgs)
        MyBase.OnLeave(e)
        If IsDate(Text) Then Text = CDate(Text).ToString("dd/MM/yyyy")
    End Sub
    Private Sub Calendar_DateSelected(sender As Object, e As DateRangeEventArgs) Handles Calendar.DateSelected
        ControlContainer.CloseDropDown()
    End Sub
    Private Sub ControlContainer_Closed(sender As Object) Handles ControlContainer.Closed
        Text = Calendar.SelectionStart.ToString("dd/MM/yyyy")
        Me.Select()
    End Sub
    Private Sub ControlContainer_Dropped(sender As Object) Handles ControlContainer.Dropped
        If Not IsDate(Text) Then
            Calendar.SetDate(Today)
        Else
            Calendar.SetDate(Text)
        End If
    End Sub
    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button.Click
        Calendar.Visible = True
    End Sub

    <DllImport("User32", SetLastError:=True)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
    End Function

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        If Not DesignMode Then
            SendMessage(Handle, &HD3, CType(2, IntPtr), CType(Button.Width << 16, IntPtr))
        End If
    End Sub

End Class
Friend Class DateBoxDesigner
    Inherits ControlDesigner
    Public Overrides Sub InitializeNewComponent(defaultValues As IDictionary)
        MyBase.InitializeNewComponent(defaultValues)
        Dim Control As DateBox = CType(Me.Control, DateBox)
        Control.Text = String.Empty
        Control.SetMask("00/00/0000")
    End Sub
    Public Overrides ReadOnly Property Verbs As DesignerVerbCollection
        Get
            Return New DesignerVerbCollection()
        End Get
    End Property
    Public Overrides ReadOnly Property ActionLists As DesignerActionListCollection
        Get
            Return New DesignerActionListCollection()
        End Get
    End Property
End Class
