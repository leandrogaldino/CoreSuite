Imports Microsoft.DotNet.DesignTools.Editors
''' <summary>
''' Provides a design-time collection editor for <see cref="NavigationPage"/> items.
''' </summary>
Public Class NavigationPageCollectionEditor
    Inherits CollectionEditor
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationPageCollectionEditor"/> class.
    ''' </summary>
    ''' <param name="ServiceProvider">The service provider supplied by the Windows Forms designer.</param>
    ''' <param name="CollectionType">The collection type edited by the designer.</param>
    Public Sub New(ServiceProvider As IServiceProvider, CollectionType As Type)
        MyBase.New(ServiceProvider, CollectionType)
    End Sub
    ''' <summary>
    ''' Gets the item types that can be created by the collection editor.
    ''' </summary>
    ''' <returns>An array containing the <see cref="NavigationPage"/> type.</returns>
    Protected Overrides Function CreateNewItemTypes() As Type()
        Return {GetType(NavigationPage)}
    End Function
End Class
