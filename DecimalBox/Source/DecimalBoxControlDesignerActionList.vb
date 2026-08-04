Imports System.ComponentModel
Imports System.Globalization
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="DecimalBox"/> control.
''' </summary>
Public Class DecimalBoxControlDesignerActionList
    Inherits DesignerActionList
    ''' <summary>
    ''' Stores the associated <see cref="DecimalBox"/> control.
    ''' </summary>
    Private ReadOnly _Control As DecimalBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DecimalBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the control.</param>
    Public Sub New(Designer As DecimalBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, DecimalBox)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart tag items displayed in the Windows Forms designer.
    ''' </summary>
    ''' <returns>A collection containing the available design-time properties.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Dim Items As New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(DecimalCulture), "DecimalCulture", "DecimalBox", "Gets or sets the culture used to parse and format values."),
            New DesignerActionPropertyItem(NameOf(DecimalValue), "DecimalValue", "DecimalBox", "Defines the decimal value represented by the control.")
        }
        Return Items
    End Function
    ''' <summary>
    ''' Gets or sets the culture used to parse and format time values.
    ''' </summary>
    Public Property DecimalCulture As CultureInfo
        Get
            Return _Control.DecimalCulture
        End Get
        Set(value As CultureInfo)
            SetProperty(NameOf(DecimalCulture), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the time represented by the associated control.
    ''' </summary>
    Public Property DecimalValue As Decimal
        Get
            Return _Control.DecimalValue
        End Get
        Set(value As Decimal)
            SetProperty(NameOf(DecimalValue), value)
        End Set
    End Property
    ''' <summary>
    ''' Sets a property on the associated control through its design-time property descriptor.
    ''' </summary>
    ''' <param name="PropertyName">The name of the property to set.</param>
    ''' <param name="Value">The value to assign to the property.</param>
    Private Sub SetProperty(PropertyName As String, Value As Object)
        TypeDescriptor.GetProperties(_Control)(PropertyName).SetValue(_Control, Value)
    End Sub
End Class

