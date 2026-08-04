''' <summary>
''' Specifies how a <see cref="DataGridViewFilterBox"/> processes filter requests.
''' </summary>
Public Enum DataGridViewFilterMode
    ''' <summary>
    ''' Applies a local filter when the target source exposes a <see cref="DataView"/>; otherwise, raises the <see cref="DataGridViewFilterBox.FilterRequested"/> event.
    ''' </summary>
    Automatic
    ''' <summary>
    ''' Requires the target source to expose a <see cref="DataView"/> and reports unsupported sources through the <see cref="DataGridViewFilterBox.FilterFailed"/> event.
    ''' </summary>
    Local
    ''' <summary>
    ''' Never changes the data source and raises the <see cref="DataGridViewFilterBox.FilterRequested"/> event so the application can perform its own filtering.
    ''' </summary>
    Custom
End Enum
