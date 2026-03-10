Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Windows.Forms.Design

<Designer(GetType(TimeBoxDesigner))>
Public Class TimeBox
    Inherits MaskedTextBox

    Private _ShowSecconds As Boolean

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    <EditorBrowsable(EditorBrowsableState.Never), Browsable(False)>
    <Description("")>
    Overloads ReadOnly Property Mask As String
        Get
            Return MyBase.Mask
        End Get
    End Property

    Friend Sub SetMask(Mask As String)
        MyBase.Mask = Mask
    End Sub

    ''' <summary>
    ''' Define a data representada pelo controle.
    ''' </summary>
    <Category("Aparência")>
    <Description("Define a data representada pelo controle.")>
    Public Property Time As TimeSpan?
        Get
            If Text.Replace(":", String.Empty).Trim = String.Empty Then Return Nothing
            Return TimeSpan.Parse(Text)
        End Get
        Set(value As TimeSpan?)
            If value IsNot Nothing Then
                Text = value.ToString
            Else
                Text = String.Empty
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define se o controle deve exibir os segundos no formato de hora.
    ''' </summary>
    ''' <value>
    ''' <c>True</c> para exibir no formato <c>HH:mm:ss</c>;
    ''' <c>False</c> para exibir apenas horas e minutos no formato <c>HH:mm</c>.
    ''' </value>
    Public Property ShowSecconds As Boolean
        Get
            Return _ShowSecconds
        End Get
        Set(value As Boolean)
            _ShowSecconds = value
            If _ShowSecconds Then
                MinimumSize = New Size(70, 0)
                MyBase.Mask = "00:00:00"
            Else
                MinimumSize = New Size(50, 0)
                MyBase.Mask = "00:00"
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define no controle apenas a parte da hora contida em um objeto <see cref="DateTime"/>.
    ''' </summary>
    ''' <param name="Date">Objeto <see cref="DateTime"/> de onde será extraída a hora.</param>
    ''' <returns>Um <see cref="TimeSpan"/> representando a hora atribuída ao controle.</returns>
    Public Function FromDateTime([Date] As Date) As TimeSpan
        Text = [Date].TimeOfDay.ToString()
    End Function

    Public Sub New()
        MyBase.Width = 50
        MyBase.TextAlign = HorizontalAlignment.Center
    End Sub

    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        BeginInvoke(CType(Sub()
                              SetMaskedTextBoxSelectAll(CType(Me, MaskedTextBox))
                          End Sub, Action))
    End Sub

    Private Sub SetMaskedTextBoxSelectAll(TextBox As MaskedTextBox)
        TextBox.SelectAll()
    End Sub

    Protected Overrides Sub OnLeave(e As EventArgs)
        MyBase.OnLeave(e)
        Dim ts As TimeSpan
        If TimeSpan.TryParse(Text, ts) Then
            Text = ts.ToString
        Else
            Text = String.Empty
        End If
    End Sub

End Class

Friend Class TimeBoxDesigner
    Inherits ControlDesigner

    Public Overrides Sub InitializeNewComponent(defaultValues As IDictionary)
        MyBase.InitializeNewComponent(defaultValues)
        Dim Control As TimeBox = CType(Me.Control, TimeBox)
        Control.Text = String.Empty
        Control.SetMask(If(Control.ShowSecconds, "00:00:00", "00:00"))
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