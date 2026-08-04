Imports System.ComponentModel
Imports System.Globalization
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="DateTimeBox"/> control.
''' </summary>
Public Class DateTimeBoxControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As DateTimeBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DateTimeBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the control.</param>
    Public Sub New(Designer As DateTimeBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, DateTimeBox)
    End Sub
    ''' <summary>
    ''' Gets the smart tag items displayed in the Windows Forms designer.
    ''' </summary>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(DateTimeCulture), "DateTimeCulture", "DateTimeBox", "Gets or sets the culture used to parse and format date and time values."),
            New DesignerActionPropertyItem(NameOf(ShowSeconds), "ShowSeconds", "DateTimeBox", "Determines whether seconds are displayed and accepted.")
        }
    End Function
    ''' <summary>
    ''' Gets or sets the culture used by the associated control.
    ''' </summary>
    Public Property DateTimeCulture As CultureInfo
        Get
            Return _Control.DateTimeCulture
        End Get
        Set(value As CultureInfo)
            SetProperty(NameOf(DateTimeCulture), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the associated control displays and accepts seconds.
    ''' </summary>
    Public Property ShowSeconds As Boolean
        Get
            Return _Control.ShowSeconds
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ShowSeconds), value)
        End Set
    End Property
    ''' <summary>
    ''' Sets a property through its design-time property descriptor.
    ''' </summary>
    ''' <param name="PropertyName">The name of the property to set.</param>
    ''' <param name="Value">The value to assign to the property.</param>
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Control)(PropertyName)
        If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        Descriptor.SetValue(_Control, Value)
    End Sub
End Class