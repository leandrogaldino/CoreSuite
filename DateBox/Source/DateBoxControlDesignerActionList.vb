Imports System.ComponentModel
Imports System.Globalization
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="DateBox"/> control.
''' </summary>
Public Class DateBoxControlDesignerActionList
    Inherits DesignerActionList
    ''' <summary>
    ''' Stores the associated <see cref="DateBox"/> control.
    ''' </summary>
    Private ReadOnly _Control As DateBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DateBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the control.</param>
    Public Sub New(Designer As DateBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, DateBox)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart tag items displayed in the Windows Forms designer.
    ''' </summary>
    ''' <returns>A collection containing the available design-time properties.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Dim Items As New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(DateCulture), "DateCulture", "DateBox", "Gets or sets the culture used to parse and format dates.")
        }
        Return Items
    End Function
    ''' <summary>
    ''' Gets or sets the culture used to parse and format date values.
    ''' </summary>
    Public Property DateCulture As CultureInfo
        Get
            Return _Control.DateCulture
        End Get
        Set(Value As CultureInfo)
            SetProperty(NameOf(DateCulture), Value)
        End Set
    End Property
    ''' <summary>
    ''' Sets a property on the associated control through its design-time property descriptor.
    ''' </summary>
    ''' <param name="PropertyName">The name of the property to set.</param>
    ''' <param name="Value">The value to assign to the property.</param>
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim PropertyDescriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Control)(PropertyName)
        If PropertyDescriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        PropertyDescriptor.SetValue(_Control, Value)
    End Sub
End Class