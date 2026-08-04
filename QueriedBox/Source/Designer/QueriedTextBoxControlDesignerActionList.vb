Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="QueriedBox"/> control.
''' </summary>
Public Class QueriedTextBoxControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As QueriedBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueriedTextBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">
    ''' The designer instance associated with the control.
    ''' </param>
    Public Sub New(Designer As QueriedTextBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, QueriedBox)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart tag items displayed in the Visual Studio designer.
    ''' </summary>
    ''' <returns>
    ''' A collection containing the available design-time actions.
    ''' </returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Dim Items As New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(QueryEnabled), "Query Enabled", "Query", "Defines whether query operations are enabled."),
            New DesignerActionPropertyItem(NameOf(QueryInterval), "Query Interval", "Query", "Defines the delay, in milliseconds, before executing a query after the text changes."),
            New DesignerActionPropertyItem(NameOf(MinimumChars), "Minimum Chars", "Query", "Defines the number of characters required to start a query.")
        }
        Return Items
    End Function
    ''' <summary>
    ''' Gets or sets whether query operations are enabled for the control.
    ''' </summary>
    Public Property QueryEnabled As Boolean
        Get
            Return _Control.Search.Enabled
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(QueryEnabled), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the interval in milliseconds between query executions.
    ''' </summary>
    Public Property QueryInterval As Integer
        Get
            Return _Control.Search.Interval
        End Get
        Set(value As Integer)
            SetProperty(NameOf(QueryInterval), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum number of characters required before starting a query.
    ''' </summary>
    Public Property MinimumChars As Integer
        Get
            Return _Control.Search.MinimumChars
        End Get
        Set(value As Integer)
            SetProperty(NameOf(MinimumChars), value)
        End Set
    End Property
    ''' <summary>
    ''' Sets a property value on the associated control through the type descriptor system.
    ''' </summary>
    ''' <param name="PropertyName">
    ''' The name of the property to update.
    ''' </param>
    ''' <param name="Value">
    ''' The value to assign to the property.
    ''' </param>
    Private Sub SetProperty(PropertyName As String, Value As Object)
        TypeDescriptor.GetProperties(_Control)(PropertyName).SetValue(_Control, Value)
    End Sub
End Class