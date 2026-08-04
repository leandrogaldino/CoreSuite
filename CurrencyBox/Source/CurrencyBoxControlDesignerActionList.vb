Imports System.ComponentModel
Imports System.Globalization
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="CurrencyBox"/> control.
''' </summary>
Public Class CurrencyBoxControlDesignerActionList
    Inherits DesignerActionList
    ''' <summary>
    ''' Stores the associated <see cref="CurrencyBox"/> control.
    ''' </summary>
    Private ReadOnly _Control As CurrencyBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CurrencyBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the control.</param>
    Public Sub New(Designer As CurrencyBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, CurrencyBox)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart tag items displayed in the Windows Forms designer.
    ''' </summary>
    ''' <returns>A collection containing the available design-time properties.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Dim Items As New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(CurrencyCulture), "CurrencyCulture", "CurrencyBox", "Gets or sets the culture used to parse and format values."),
            New DesignerActionPropertyItem(NameOf(CurrencyValue), "CurrencyValue", "CurrencyBox", "Defines the value represented by the control.")
        }
        Return Items
    End Function
    ''' <summary>
    ''' Gets or sets the culture used to parse and format time values.
    ''' </summary>
    Public Property CurrencyCulture As CultureInfo
        Get
            Return _Control.CurrencyCulture
        End Get
        Set(value As CultureInfo)
            SetProperty(NameOf(CurrencyCulture), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the time represented by the associated control.
    ''' </summary>
    Public Property CurrencyValue As Decimal
        Get
            Return _Control.CurrencyValue
        End Get
        Set(value As Decimal)
            SetProperty(NameOf(CurrencyValue), value)
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

