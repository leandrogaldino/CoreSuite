Imports System.ComponentModel
Imports System.Data.Common
Partial Public Class QueriedBox
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property Multiline As Boolean
        Get
            Return MyBase.Multiline
        End Get
        Set(value As Boolean)
            If value Then
                Search.Enabled = False
            End If
            MyBase.Multiline = value
        End Set
    End Property
    Public Overrides Property ForeColor As Color
        Get
            Return MyBase.ForeColor
        End Get
        Set(value As Color)
            MyBase.ForeColor = value
            If Not _Freezing Then
                Frozen.SetUnFrozenColor(value)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Defines the factory method used to create database connections for query execution.
    ''' </summary>
    ''' <remarks>
    ''' The factory should return a new database connection instance whenever invoked.
    ''' The lifetime and disposal of the created connection are managed internally by the control.
    ''' </remarks>
    <Browsable(False)>
    Public Shared Property ConnectionFactory As Func(Of DbConnection)
    ''' <summary>
    ''' Gets or sets the default SQL dialect used by queries that do not specify a dialect explicitly.
    ''' </summary>
    ''' <remarks>
    ''' This value is used when <see cref="Query.Dialect"/> is <see langword="Nothing"/>.
    ''' Individual queries can override this value by setting their own dialect.
    ''' </remarks>
    Public Shared Property SqlDialect As SqlDialect = SqlDialect.MySql
    ''' <summary>
    ''' Gets or sets the query configuration used by the control.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the query configuration used by the control.")>
    Public Property Query As New Query()
    ''' <summary>
    ''' Gets or sets the search behavior settings used by the control.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the search behavior settings used by the control.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Search As New QueriedBoxSearchOptions()
    ''' <summary>
    ''' Gets or sets the drop-down window settings used by the control.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the drop-down window settings used by the control.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property DropDown As New QueriedBoxDropDownOptions()
    ''' <summary>
    ''' Gets or sets the result grid appearance and behavior settings.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the result grid appearance and behavior settings.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Grid As New QueriedBoxGridOptions()
    ''' <summary>
    ''' Gets or sets the messages and message appearance used by the control.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the messages and message appearance used by the control.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Messages As New QueriedBoxMessageOptions()
    ''' <summary>
    ''' Gets or sets the appearance and behavior applied to frozen values.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the appearance and behavior applied to frozen values.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Frozen As New QueriedBoxFrozenOptions()
    ''' <summary>
    ''' Gets or sets the diagnostic settings used by the control.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the diagnostic settings used by the control.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Diagnostics As New QueriedBoxDiagnosticsOptions()
    ''' <summary>
    ''' Gets or sets whether the beginning of the content is displayed when leaving the control.
    ''' </summary>
    <Category("QueriedBox")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Gets or sets whether the beginning of the content is displayed when leaving the control.")>
    Public Property SelectFromStartOnLeave As Boolean = True
End Class
