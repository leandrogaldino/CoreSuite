Imports System.ComponentModel.Design
''' <summary>
''' Provides a design-time collection editor for <see cref="LookupColumn"/> items.
''' </summary>
Public Class LookupColumnCollectionEditor
    Inherits CollectionEditor
    ''' <summary>
    ''' Initializes a new instance of the <see cref="LookupColumnCollectionEditor"/> class.
    ''' </summary>
    ''' <param name="CollectionType">The collection type edited by the designer.</param>
    Public Sub New(CollectionType As Type)
        MyBase.New(CollectionType)
    End Sub
    Public Sub New(ServiceProvider As IServiceProvider, CollectionType As Type)
        MyBase.New(CollectionType)
    End Sub
    ''' <summary>
    ''' Gets the item types that can be created by the collection editor.
    ''' </summary>
    ''' <returns>An array containing the <see cref="LookupColumn"/> type.</returns>
    Protected Overrides Function CreateNewItemTypes() As Type()
        Return {GetType(LookupColumn)}
    End Function
End Class
