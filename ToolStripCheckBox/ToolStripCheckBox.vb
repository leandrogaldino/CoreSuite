Imports System.Windows.Forms
Imports System.Windows.Forms.Design
''' <summary>
''' Represents a ToolStrip item that hosts a CheckBox control.
''' Allows placing a standard CheckBox inside a ToolStrip.
''' </summary>
<System.ComponentModel.DesignerCategory("code")>
<ToolStripItemDesignerAvailability(ToolStripItemDesignerAvailability.ToolStrip)>
Public Class ToolStripCheckBox
    Inherits ToolStripControlHost
    ''' <summary>
    ''' Occurs when the CheckBox is checked or unchecked.
    ''' </summary>
    Public Event CheckedChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ToolStripCheckBox"/> class.
    ''' </summary>
    Public Sub New()
        MyBase.New(New CheckBox())
    End Sub
    Public ReadOnly Property CheckBoxControl As CheckBox
        Get
            Return TryCast(Control, CheckBox)
        End Get
    End Property
    ''' <summary>
    ''' Gets the hosted CheckBox control.
    ''' </summary>
    Public Property Checked As Boolean
        Get
            Return CheckBoxControl.Checked
        End Get
        Set(ByVal value As Boolean)
            CheckBoxControl.Checked = value
        End Set
    End Property
    Protected Overrides Sub OnSubscribeControlEvents(ByVal c As Control)
        MyBase.OnSubscribeControlEvents(c)
        Dim CheckBoxControl As CheckBox = CType(c, CheckBox)
        AddHandler CheckBoxControl.CheckedChanged, New EventHandler(AddressOf OnCheckedChanged)
    End Sub
    Protected Overrides Sub OnUnsubscribeControlEvents(ByVal c As Control)
        MyBase.OnUnsubscribeControlEvents(c)
        Dim CheckBoxControl As CheckBox = CType(c, CheckBox)
        RemoveHandler CheckBoxControl.CheckedChanged, New EventHandler(AddressOf OnCheckedChanged)
        AddHandler CheckBoxControl.MouseHover, New EventHandler(AddressOf Me_MouseHover)
    End Sub
    Private Sub OnCheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent CheckedChanged(Me, e)
    End Sub
    Private Sub Me_MouseHover(ByVal sender As Object, ByVal e As EventArgs)
        Me.OnMouseHover(e)
    End Sub
End Class