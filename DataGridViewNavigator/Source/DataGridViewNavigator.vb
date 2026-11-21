Imports System.ComponentModel
Imports System.Windows.Forms

''' <summary>
''' Provides first, previous, next and last row navigation for a DataGridView.
''' </summary>
''' <remarks>
''' Assign the grid and the desired ToolStrip items in the designer or in code.
''' Attaching a grid sets MultiSelect=False and SelectionMode=FullRowSelect.
''' All public members must be used on the WinForms UI thread.
''' </remarks>
<DesignerCategory("Component")>
Public Partial Class DataGridViewNavigator
    Inherits Component

    Private _DataGridView As DataGridView
    Private _FirstButton As ToolStripButton
    Private _PreviousButton As ToolStripButton
    Private _NextButton As ToolStripButton
    Private _LastButton As ToolStripButton
    Private _PositionLabel As ToolStripLabel
    Private _PositionFormat As String = "{0} of {1}"
    Private _IsNavigating As Boolean
    Private _IsDisposed As Boolean

    ' Pending selection recovery following a row removal.
    Private _PendingSelectionRestoreIndex As Integer = -1
    Private _SelectionRestoreQueued As Boolean
    Private _GridGeneration As Integer

    Public Sub New()
    End Sub

    ''' <summary>Creates the component inside a design-time or runtime container.</summary>
    Public Sub New(Container As IContainer)
        Me.New()
        If Container Is Nothing Then Throw New ArgumentNullException(NameOf(Container))
        Container.Add(Me)
    End Sub

    ''' <summary>Runs before navigation; set CancelEventArgs.Cancel to cancel.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ActionBeforeMove As Action(Of CancelEventArgs)

    ''' <summary>Runs asynchronously before navigation; may cancel through CancelEventArgs.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ActionBeforeMoveAsync As Func(Of CancelEventArgs, Task)

    ''' <summary>Runs after the selection has been changed successfully.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ActionAfterMove As Action

    ''' <summary>Runs asynchronously after the selection has been changed successfully.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ActionAfterMoveAsync As Func(Of Task)

    ''' <summary>True while a navigation operation is in progress.</summary>
    <Browsable(False)>
    Public ReadOnly Property IsNavigating As Boolean
        Get
            Return _IsNavigating
        End Get
    End Property

    ''' <summary>The grid navigated by this component.</summary>
    <Category("Navigation")>
    <Description("DataGridView used for record navigation.")>
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            ThrowIfDisposed()
            If ReferenceEquals(_DataGridView, value) Then Return
            If value IsNot Nothing AndAlso (value.IsDisposed OrElse value.Disposing) Then
                Throw New ObjectDisposedException(NameOf(value))
            End If

            DetachGrid(_DataGridView)
            _GridGeneration += 1
            _PendingSelectionRestoreIndex = -1
            _SelectionRestoreQueued = False
            _DataGridView = Nothing

            If value IsNot Nothing Then
                value.MultiSelect = False
                value.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                _DataGridView = value
                AttachGrid(value)
            End If

            RefreshButtons()
        End Set
    End Property

    ''' <summary>Format string with {0}=position and {1}=total navigable rows.</summary>
    <Category("Navigation")>
    <Description("Format used for the current row and the total number of rows.")>
    <DefaultValue("{0} of {1}")>
    Public Property PositionFormat As String
        Get
            Return _PositionFormat
        End Get
        Set(value As String)
            ThrowIfDisposed()
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            ' Validate the format at assignment, not during a grid event.
            String.Format(value, 0, 0)
            _PositionFormat = value
            RefreshButtons()
        End Set
    End Property

    ''' <summary>Optional ToolStripLabel that displays the current position.</summary>
    <Category("Navigation")>
    <Description("ToolStripLabel displaying the current position and total.")>
    Public Property PositionLabel As ToolStripLabel
        Get
            Return _PositionLabel
        End Get
        Set(value As ToolStripLabel)
            ThrowIfDisposed()
            If ReferenceEquals(_PositionLabel, value) Then Return
            _PositionLabel = value
            RefreshButtons()
        End Set
    End Property

    <Category("Navigation")>
    <Description("ToolStripButton that navigates to the first row.")>
    Public Property FirstButton As ToolStripButton
        Get
            Return _FirstButton
        End Get
        Set(value As ToolStripButton)
            ThrowIfDisposed()
            If ReferenceEquals(_FirstButton, value) Then Return
            If _FirstButton IsNot Nothing Then RemoveHandler _FirstButton.Click, AddressOf FirstButton_Click
            _FirstButton = value
            If _FirstButton IsNot Nothing Then AddHandler _FirstButton.Click, AddressOf FirstButton_Click
            RefreshButtons()
        End Set
    End Property

    <Category("Navigation")>
    <Description("ToolStripButton that navigates to the previous row.")>
    Public Property PreviousButton As ToolStripButton
        Get
            Return _PreviousButton
        End Get
        Set(value As ToolStripButton)
            ThrowIfDisposed()
            If ReferenceEquals(_PreviousButton, value) Then Return
            If _PreviousButton IsNot Nothing Then RemoveHandler _PreviousButton.Click, AddressOf PreviousButton_Click
            _PreviousButton = value
            If _PreviousButton IsNot Nothing Then AddHandler _PreviousButton.Click, AddressOf PreviousButton_Click
            RefreshButtons()
        End Set
    End Property

    <Category("Navigation")>
    <Description("ToolStripButton that navigates to the next row.")>
    Public Property NextButton As ToolStripButton
        Get
            Return _NextButton
        End Get
        Set(value As ToolStripButton)
            ThrowIfDisposed()
            If ReferenceEquals(_NextButton, value) Then Return
            If _NextButton IsNot Nothing Then RemoveHandler _NextButton.Click, AddressOf NextButton_Click
            _NextButton = value
            If _NextButton IsNot Nothing Then AddHandler _NextButton.Click, AddressOf NextButton_Click
            RefreshButtons()
        End Set
    End Property

    <Category("Navigation")>
    <Description("ToolStripButton that navigates to the last row.")>
    Public Property LastButton As ToolStripButton
        Get
            Return _LastButton
        End Get
        Set(value As ToolStripButton)
            ThrowIfDisposed()
            If ReferenceEquals(_LastButton, value) Then Return
            If _LastButton IsNot Nothing Then RemoveHandler _LastButton.Click, AddressOf LastButton_Click
            _LastButton = value
            If _LastButton IsNot Nothing Then AddHandler _LastButton.Click, AddressOf LastButton_Click
            RefreshButtons()
        End Set
    End Property

    Private Sub ThrowIfDisposed()
        If _IsDisposed Then Throw New ObjectDisposedException(NameOf(DataGridViewNavigator))
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not _IsDisposed Then
            _IsDisposed = True
            DetachGrid(_DataGridView)
            _DataGridView = Nothing

            If _FirstButton IsNot Nothing Then RemoveHandler _FirstButton.Click, AddressOf FirstButton_Click
            If _PreviousButton IsNot Nothing Then RemoveHandler _PreviousButton.Click, AddressOf PreviousButton_Click
            If _NextButton IsNot Nothing Then RemoveHandler _NextButton.Click, AddressOf NextButton_Click
            If _LastButton IsNot Nothing Then RemoveHandler _LastButton.Click, AddressOf LastButton_Click

            _FirstButton = Nothing
            _PreviousButton = Nothing
            _NextButton = Nothing
            _LastButton = Nothing
            _PositionLabel = Nothing
        End If

        MyBase.Dispose(disposing)
    End Sub
End Class
