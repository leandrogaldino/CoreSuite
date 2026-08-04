Imports System.ComponentModel
''' <summary>
''' Provides the tooltip texts displayed by the <see cref="PictureViewer"/> toolbar buttons.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class PictureViewerToolTips
    Friend Event Changed()
    Private _First As String = "First"
    Private _Previous As String = "Previous"
    Private _Next As String = "Next"
    Private _Last As String = "Last"
    Private _Include As String = "Include"
    Private _Remove As String = "Remove"
    Private _Save As String = "Save"
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the First button.
    ''' </summary>
    <DefaultValue("First")>
    Public Property First As String
        Get
            Return _First
        End Get
        Set(value As String)
            If _First = value Then Return
            _First = value
            RaiseEvent Changed()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the Previous button.
    ''' </summary>
    <DefaultValue("Previous")>
    Public Property Previous As String
        Get
            Return _Previous
        End Get
        Set(value As String)
            If _Previous = value Then Return
            _Previous = value
            RaiseEvent Changed()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the Next button.
    ''' </summary>
    <DefaultValue("Next")>
    Public Property [Next] As String
        Get
            Return _Next
        End Get
        Set(value As String)
            If _Next = value Then Return
            _Next = value
            RaiseEvent Changed()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the Last button.
    ''' </summary>
    <DefaultValue("Last")>
    Public Property Last As String
        Get
            Return _Last
        End Get
        Set(value As String)
            If _Last = value Then Return
            _Last = value
            RaiseEvent Changed()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the Include button.
    ''' </summary>
    <DefaultValue("Include")>
    Public Property Include As String
        Get
            Return _Include
        End Get
        Set(value As String)
            If _Include = value Then Return
            _Include = value
            RaiseEvent Changed()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the Remove button.
    ''' </summary>
    <DefaultValue("Remove")>
    Public Property Remove As String
        Get
            Return _Remove
        End Get
        Set(value As String)
            If _Remove = value Then Return
            _Remove = value
            RaiseEvent Changed()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the Save button.
    ''' </summary>
    <DefaultValue("Save")>
    Public Property Save As String
        Get
            Return _Save
        End Get
        Set(value As String)
            If _Save = value Then Return
            _Save = value
            RaiseEvent Changed()
        End Set
    End Property
End Class