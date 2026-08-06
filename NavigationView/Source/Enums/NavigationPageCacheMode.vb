''' <summary>
''' Specifies how a <see cref="NavigationPage"/> manages the lifetime of its created <see cref="UserControl"/>.
''' </summary>
Public Enum NavigationPageCacheMode
    ''' <summary>
    ''' Keeps the created control in memory and preserves its state while another page is displayed.
    ''' </summary>
    KeepAlive
    ''' <summary>
    ''' Disposes the created control when navigation leaves the page and creates a new instance when the page is opened again.
    ''' </summary>
    Recreate
End Enum
