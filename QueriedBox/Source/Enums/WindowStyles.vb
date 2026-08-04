''' <summary>
''' Specifies window styles used when creating or configuring a native Windows window.
''' This enumeration supports combining multiple values using bitwise operations.
''' </summary>
<Flags()>
Public Enum WindowStyles As UInteger
    ''' <summary>
    ''' Creates an overlapped window.
    ''' </summary>
    WS_OVERLAPPED = 0
    ''' <summary>
    ''' Creates a pop-up window.
    ''' </summary>
    WS_POPUP = 2147483648
    ''' <summary>
    ''' Creates a child window.
    ''' </summary>
    WS_CHILD = 1073741824
    ''' <summary>
    ''' Creates a minimized window.
    ''' </summary>
    WS_MINIMIZE = 536870912
    ''' <summary>
    ''' Creates a visible window.
    ''' </summary>
    WS_VISIBLE = 268435456
    ''' <summary>
    ''' Creates a disabled window.
    ''' </summary>
    WS_DISABLED = 134217728
    ''' <summary>
    ''' Clips child windows relative to each other.
    ''' </summary>
    WS_CLIPSIBLINGS = 67108864
    ''' <summary>
    ''' Prevents drawing over child window areas.
    ''' </summary>
    WS_CLIPCHILDREN = 33554432
    ''' <summary>
    ''' Creates a maximized window.
    ''' </summary>
    WS_MAXIMIZE = 16777216
    ''' <summary>
    ''' Creates a window with a border.
    ''' </summary>
    WS_BORDER = 8388608
    ''' <summary>
    ''' Creates a window with a dialog frame.
    ''' </summary>
    WS_DLGFRAME = 4194304
    ''' <summary>
    ''' Adds a vertical scroll bar.
    ''' </summary>
    WS_VSCROLL = 2097152
    ''' <summary>
    ''' Adds a horizontal scroll bar.
    ''' </summary>
    WS_HSCROLL = 1048576
    ''' <summary>
    ''' Adds a system menu to the window.
    ''' </summary>
    WS_SYSMENU = 524288
    ''' <summary>
    ''' Creates a resizable window.
    ''' </summary>
    WS_THICKFRAME = 262144
    ''' <summary>
    ''' Defines a control group.
    ''' </summary>
    WS_GROUP = 131072
    ''' <summary>
    ''' Allows keyboard navigation using the TAB key.
    ''' </summary>
    WS_TABSTOP = 65536
    ''' <summary>
    ''' Creates a minimize button.
    ''' </summary>
    WS_MINIMIZEBOX = 131072
    ''' <summary>
    ''' Creates a maximize button.
    ''' </summary>
    WS_MAXIMIZEBOX = 65536
    ''' <summary>
    ''' Creates a window caption.
    ''' </summary>
    WS_CAPTION = WS_BORDER Or WS_DLGFRAME
    ''' <summary>
    ''' Alias for WS_OVERLAPPED.
    ''' </summary>
    WS_TILED = WS_OVERLAPPED
    ''' <summary>
    ''' Alias for WS_MINIMIZE.
    ''' </summary>
    WS_ICONIC = WS_MINIMIZE
    ''' <summary>
    ''' Alias for WS_THICKFRAME.
    ''' </summary>
    WS_SIZEBOX = WS_THICKFRAME
    ''' <summary>
    ''' Creates a standard overlapped window.
    ''' </summary>
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    ''' <summary>
    ''' Creates a standard overlapped window with common window components.
    ''' </summary>
    WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    ''' <summary>
    ''' Creates a pop-up window with a border and system menu.
    ''' </summary>
    WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
    ''' <summary>
    ''' Alias for WS_CHILD.
    ''' </summary>
    WS_CHILDWINDOW = WS_CHILD
    ''' <summary>
    ''' Creates a dialog modal frame.
    ''' </summary>
    WS_EX_DLGMODALFRAME = 1
    ''' <summary>
    ''' Prevents notification messages from being sent to the parent window.
    ''' </summary>
    WS_EX_NOPARENTNOTIFY = 4
    ''' <summary>
    ''' Places the window above all non-topmost windows.
    ''' </summary>
    WS_EX_TOPMOST = 8
    ''' <summary>
    ''' Allows the window to accept dropped files.
    ''' </summary>
    WS_EX_ACCEPTFILES = 16
    ''' <summary>
    ''' Creates a transparent window.
    ''' </summary>
    WS_EX_TRANSPARENT = 32
    ''' <summary>
    ''' Creates an MDI child window.
    ''' </summary>
    WS_EX_MDICHILD = 64
    ''' <summary>
    ''' Creates a tool window.
    ''' </summary>
    WS_EX_TOOLWINDOW = 128
    ''' <summary>
    ''' Adds a raised edge border.
    ''' </summary>
    WS_EX_WINDOWEDGE = 256
    ''' <summary>
    ''' Adds a sunken client edge border.
    ''' </summary>
    WS_EX_CLIENTEDGE = 512
    ''' <summary>
    ''' Adds a context help button to the title bar.
    ''' </summary>
    WS_EX_CONTEXTHELP = 1024
    ''' <summary>
    ''' Creates a right-aligned window.
    ''' </summary>
    WS_EX_RIGHT = 4096
    ''' <summary>
    ''' Creates a left-aligned window.
    ''' </summary>
    WS_EX_LEFT = 0
    ''' <summary>
    ''' Enables right-to-left reading order.
    ''' </summary>
    WS_EX_RTLREADING = 8192
    ''' <summary>
    ''' Enables left-to-right reading order.
    ''' </summary>
    WS_EX_LTRREADING = 0
    ''' <summary>
    ''' Places the vertical scroll bar on the left side.
    ''' </summary>
    WS_EX_LEFTSCROLLBAR = 16384
    ''' <summary>
    ''' Places the vertical scroll bar on the right side.
    ''' </summary>
    WS_EX_RIGHTSCROLLBAR = 0
    ''' <summary>
    ''' Allows child controls to participate in dialog navigation.
    ''' </summary>
    WS_EX_CONTROLPARENT = 65536
    ''' <summary>
    ''' Creates a static edge border.
    ''' </summary>
    WS_EX_STATICEDGE = 131072
    ''' <summary>
    ''' Forces a top-level window to appear in the taskbar.
    ''' </summary>
    WS_EX_APPWINDOW = 262144
    ''' <summary>
    ''' Creates a window with window and client edges.
    ''' </summary>
    WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
    ''' <summary>
    ''' Creates a palette window style.
    ''' </summary>
    WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
    ''' <summary>
    ''' Enables layered window rendering.
    ''' </summary>
    WS_EX_LAYERED = 524288
    ''' <summary>
    ''' Prevents layout inheritance.
    ''' </summary>
    WS_EX_NOINHERITLAYOUT = 1048576
    ''' <summary>
    ''' Enables right-to-left layout.
    ''' </summary>
    WS_EX_LAYOUTRTL = 4194304
    ''' <summary>
    ''' Enables composited drawing to reduce flickering.
    ''' </summary>
    WS_EX_COMPOSITED = 33554432
    ''' <summary>
    ''' Prevents window activation when displayed.
    ''' </summary>
    WS_EX_NOACTIVATE = 134217728
End Enum