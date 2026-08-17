Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="ActionTextBox"/> control.
''' </summary>
Public Class ActionTextBoxControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As ActionTextBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ActionTextBoxControlDesignerActionList"/> class.
    ''' </summary>
    Public Sub New(designer As ActionTextBoxControlDesigner)
        MyBase.New(designer.Component)
        _Control = CType(designer.Component, ActionTextBox)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart tag items displayed in the Windows Forms designer.
    ''' </summary>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(ActionButtonImage), "ActionButtonImage", "Action Button", "Specifies the image displayed by the action button."),
            New DesignerActionPropertyItem(NameOf(ActionButtonPosition), "ActionButtonPosition", "Action Button", "Specifies the side on which the action button is displayed."),
            New DesignerActionPropertyItem(NameOf(ActionButtonVisible), "ActionButtonVisible", "Action Button", "Specifies whether the action button is visible."),
            New DesignerActionPropertyItem(NameOf(ActionButtonEnabled), "ActionButtonEnabled", "Action Button", "Specifies whether the action button can be clicked."),
            New DesignerActionPropertyItem(NameOf(ActionButtonWidth), "ActionButtonWidth", "Action Button", "Specifies the width of the action button."),
            New DesignerActionPropertyItem(NameOf(ActionButtonPadding), "ActionButtonPadding", "Action Button", "Specifies the padding around the action button image."),
            New DesignerActionPropertyItem(NameOf(ActionButtonTextSpacing), "ActionButtonTextSpacing", "Action Button", "Specifies the spacing between the text and the action button in logical pixels."),
            New DesignerActionPropertyItem(NameOf(ActionButtonToolTipText), "ActionButtonToolTipText", "Action Button", "Specifies the tooltip text displayed for the action button.")
        }
    End Function
    ''' <summary>
    ''' Gets or sets the image displayed by the action button.
    ''' </summary>
    Public Property ActionButtonImage As Image
        Get
            Return _Control.ActionButtonImage
        End Get
        Set(value As Image)
            SetProperty(NameOf(ActionButtonImage), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the side on which the action button is displayed.
    ''' </summary>
    Public Property ActionButtonPosition As ActionButtonPosition
        Get
            Return _Control.ActionButtonPosition
        End Get
        Set(value As ActionButtonPosition)
            SetProperty(NameOf(ActionButtonPosition), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the action button is visible.
    ''' </summary>
    Public Property ActionButtonVisible As Boolean
        Get
            Return _Control.ActionButtonVisible
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ActionButtonVisible), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the action button can be clicked.
    ''' </summary>
    Public Property ActionButtonEnabled As Boolean
        Get
            Return _Control.ActionButtonEnabled
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ActionButtonEnabled), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the width of the action button.
    ''' </summary>
    Public Property ActionButtonWidth As Integer
        Get
            Return _Control.ActionButtonWidth
        End Get
        Set(value As Integer)
            SetProperty(NameOf(ActionButtonWidth), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the padding around the action button image.
    ''' </summary>
    Public Property ActionButtonPadding As Integer
        Get
            Return _Control.ActionButtonPadding
        End Get
        Set(value As Integer)
            SetProperty(NameOf(ActionButtonPadding), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the action button.
    ''' </summary>
    Public Property ActionButtonToolTipText As String
        Get
            Return _Control.ActionButtonToolTipText
        End Get
        Set(value As String)
            SetProperty(NameOf(ActionButtonToolTipText), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the spacing between the text and the action button in logical pixels.
    ''' </summary>
    Public Property ActionButtonTextSpacing As Integer
        Get
            Return _Control.ActionButtonTextSpacing
        End Get
        Set(value As Integer)
            SetProperty(NameOf(ActionButtonTextSpacing), value)
        End Set
    End Property

    Private Sub SetProperty(propertyName As String, value As Object)
        TypeDescriptor.GetProperties(_Control)(propertyName).SetValue(_Control, value)
    End Sub
End Class
