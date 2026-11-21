Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports System.Drawing.ContentAlignment
Imports System.Drawing

''' <summary>
''' Represents a ToolStrip item that hosts a standard CheckBox control.
''' </summary>
''' <remarks>
''' Provides direct access to common CheckBox properties and events
''' while preserving the native Windows Forms appearance and behavior.
''' </remarks>
<DesignerCategory("Code")>
<ToolStripItemDesignerAvailability(ToolStripItemDesignerAvailability.ToolStrip)>
Public Class ToolStripCheckBox
    Inherits ToolStripControlHost



    ''' <summary>
    ''' Occurs when the checked state changes.
    ''' </summary>
    <Category("Action")>
    <Description("Occurs when the checked state changes.")>
    Public Event CheckedChanged As EventHandler

    ''' <summary>
    ''' Occurs when the CheckState property changes.
    ''' </summary>
    <Category("Action")>
    <Description("Occurs when the check state changes.")>
    Public Event CheckStateChanged As EventHandler



    ''' <summary>
    ''' Initializes a new instance of the ToolStripCheckBox class.
    ''' </summary>
    Public Sub New()
        MyBase.New(New CheckBox With {
            .AutoSize = True,
            .Text = String.Empty
        })
    End Sub



    ''' <summary>
    ''' Gets the hosted CheckBox control.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property CheckBoxControl As CheckBox
        Get
            Return DirectCast(Control, CheckBox)
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets whether the CheckBox is checked.
    ''' </summary>
    <Category("Behavior")>
    <Description("Indicates whether the CheckBox is checked.")>
    <DefaultValue(False)>
    Public Property Checked As Boolean
        Get
            Return CheckBoxControl.Checked
        End Get
        Set(Value As Boolean)
            CheckBoxControl.Checked = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the current check state.
    ''' </summary>
    <Category("Behavior")>
    <Description("Specifies the current check state.")>
    <DefaultValue(CheckState.Unchecked)>
    Public Property CheckState As CheckState
        Get
            Return CheckBoxControl.CheckState
        End Get
        Set(Value As CheckState)
            CheckBoxControl.CheckState = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether the CheckBox supports three states.
    ''' </summary>
    <Category("Behavior")>
    <Description("Indicates whether the CheckBox supports three states.")>
    <DefaultValue(False)>
    Public Property ThreeState As Boolean
        Get
            Return CheckBoxControl.ThreeState
        End Get
        Set(Value As Boolean)
            CheckBoxControl.ThreeState = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets whether the check state changes automatically
    ''' when the user clicks the CheckBox.
    ''' </summary>
    <Category("Behavior")>
    <Description("Indicates whether clicking automatically changes the check state.")>
    <DefaultValue(True)>
    Public Property AutoCheck As Boolean
        Get
            Return CheckBoxControl.AutoCheck
        End Get
        Set(Value As Boolean)
            CheckBoxControl.AutoCheck = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the appearance of the hosted CheckBox.
    ''' </summary>
    <Category("Appearance")>
    <Description("Specifies whether the CheckBox appears as a standard checkbox or a button.")>
    <DefaultValue(Appearance.Normal)>
    Public Property Appearance As Appearance
        Get
            Return CheckBoxControl.Appearance
        End Get
        Set(Value As Appearance)
            CheckBoxControl.Appearance = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the alignment of the check indicator.
    ''' </summary>
    <Category("Appearance")>
    <Description("Specifies the alignment of the check indicator.")>
    <DefaultValue(ContentAlignment.MiddleLeft)>
    Public Property CheckAlign As ContentAlignment
        Get
            Return CheckBoxControl.CheckAlign
        End Get
        Set(Value As ContentAlignment)
            CheckBoxControl.CheckAlign = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the alignment of the CheckBox text.
    ''' </summary>
    ''' <remarks>
    ''' Hides the ToolStripControlHost TextAlign property to expose
    ''' the alignment of the actual hosted CheckBox.
    ''' </remarks>
    <Browsable(True)>
    <Category("Appearance")>
    <Description("Specifies the alignment of the CheckBox text.")>
    <DefaultValue(ContentAlignment.MiddleLeft)>
    Public Shadows Property TextAlign As ContentAlignment
        Get
            Return CheckBoxControl.TextAlign
        End Get
        Set(Value As ContentAlignment)
            CheckBoxControl.TextAlign = Value
        End Set
    End Property




    ''' <summary>
    ''' Subscribes to events raised by the hosted CheckBox.
    ''' </summary>
    Protected Overrides Sub OnSubscribeControlEvents(C As Control)
        MyBase.OnSubscribeControlEvents(C)

        Dim HostedCheckBox As CheckBox = DirectCast(C, CheckBox)

        AddHandler HostedCheckBox.CheckedChanged, AddressOf OnHostedCheckedChanged
        AddHandler HostedCheckBox.CheckStateChanged, AddressOf OnHostedCheckStateChanged
    End Sub

    ''' <summary>
    ''' Unsubscribes from events raised by the hosted CheckBox.
    ''' </summary>
    Protected Overrides Sub OnUnsubscribeControlEvents(C As Control)
        Dim HostedCheckBox As CheckBox = DirectCast(C, CheckBox)

        RemoveHandler HostedCheckBox.CheckedChanged, AddressOf OnHostedCheckedChanged
        RemoveHandler HostedCheckBox.CheckStateChanged, AddressOf OnHostedCheckStateChanged

        MyBase.OnUnsubscribeControlEvents(C)
    End Sub

    ''' <summary>
    ''' Raises the CheckedChanged event.
    ''' </summary>
    Protected Overridable Sub OnCheckedChanged(E As EventArgs)
        RaiseEvent CheckedChanged(Me, E)
    End Sub

    ''' <summary>
    ''' Raises the CheckStateChanged event.
    ''' </summary>
    Protected Overridable Sub OnCheckStateChanged(E As EventArgs)
        RaiseEvent CheckStateChanged(Me, E)
    End Sub



    ''' <summary>
    ''' Handles changes to the hosted CheckBox checked state.
    ''' </summary>
    Private Sub OnHostedCheckedChanged(Sender As Object, E As EventArgs)
        OnCheckedChanged(E)
    End Sub

    ''' <summary>
    ''' Handles changes to the hosted CheckBox check state.
    ''' </summary>
    Private Sub OnHostedCheckStateChanged(Sender As Object, E As EventArgs)
        OnCheckStateChanged(E)
    End Sub


End Class
