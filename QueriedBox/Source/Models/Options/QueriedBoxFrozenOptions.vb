Imports System.ComponentModel
''' <summary>
''' Provides configuration options related to frozen values in a <see cref="QueriedBox"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxFrozenOptions
    Private _FrozenColor As Color = Color.Blue
    Private _UnFrozenColor As Color
    Private _IsFrozen As Boolean
    Private _FrozenPrimaryKey As Object
    Private _FrozenValue As String
    ''' <summary>
    ''' Gets or sets the text color used when a query result is selected.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Color), "Blue")>
    <Description("Gets or sets the text color used when a query result is selected.")>
    Public Property FrozenColor As Color
        Get
            Return _FrozenColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _FrozenColor = value
            Else
                Common.ThrowNoTransparentColorException()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the control clears its content when a value is unfrozen.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Gets or sets whether the control clears its content when a value is unfrozen.")>
    Public Property ClearOnUnfreeze As Boolean
    ''' <summary>
    ''' Gets or sets whether hyperlinks are enabled for selected records when using the control key.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Gets or sets whether hyperlinks are enabled for selected records when using the control key.")>
    Public Property AllowHyperlink As Boolean = False
    ''' <summary>
    ''' Gets or sets whether the beginning of the content is displayed when freezing a result.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Gets or sets whether the beginning of the content is displayed when freezing a result.")>
    Public Property ShowStartOnFreeze As Boolean = False
    ''' <summary>
    ''' Gets whether a record is currently selected.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsFrozen As Boolean
        Get
            Return _IsFrozen
        End Get
    End Property
    ''' <summary>
    ''' Gets the primary key value of the selected record.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property FrozenPrimaryKey As Object
        Get
            Return _FrozenPrimaryKey
        End Get
    End Property
    ''' <summary>
    ''' Gets the displayed value of the selected record.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property FrozenValue As String
        Get
            Return _FrozenValue
        End Get
    End Property
    ''' <summary>
    ''' Gets the default color used when the control is not frozen.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property UnFrozenColor As Color
        Get
            Return _UnFrozenColor
        End Get
    End Property
    ''' <summary>
    ''' Updates the default text color used when the control is not frozen.
    ''' </summary>
    ''' <param name="Color">The text color to apply while the control is not frozen.</param>
    Friend Sub SetUnFrozenColor(Color As Color)
        _UnFrozenColor = Color
    End Sub
    ''' <summary>
    ''' Updates the display value of the currently frozen record.
    ''' </summary>
    ''' <param name="Value">The display value to associate with the frozen record.</param>
    Friend Sub SetFrozenValue(Value As String)
        _FrozenValue = Value
    End Sub
    ''' <summary>
    ''' Updates the primary key of the currently frozen record.
    ''' </summary>
    ''' <param name="PrimaryKey">The primary key value to associate with the frozen record.</param>
    Friend Sub SetFrozenPrimaryKey(PrimaryKey As Object)
        _FrozenPrimaryKey = PrimaryKey
    End Sub
    ''' <summary>
    ''' Updates the frozen state of the control.
    ''' </summary>
    ''' <param name="IsFrozen"><see langword="True"/> to mark the control as frozen; otherwise, <see langword="False"/>.</param>
    Friend Sub SetIsFrozen(IsFrozen As Boolean)
        _IsFrozen = IsFrozen
    End Sub
    ''' <summary>
    ''' Returns a summary of the configured frozen value options.
    ''' </summary>
    ''' <returns>
    ''' A string describing the frozen color, clearing behavior, hyperlink support,
    ''' and content positioning.
    ''' </returns>
    Public Overrides Function ToString() As String
        Return $"Color: {FrozenColor.Name}, " &
               $"Clear on unfreeze: {If(ClearOnUnfreeze, "Yes", "No")}, " &
               $"Hyperlink: {If(AllowHyperlink, "Enabled", "Disabled")}, " &
               $"Show start: {If(ShowStartOnFreeze, "Yes", "No")}"
    End Function
End Class
