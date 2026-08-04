Imports System.ComponentModel
''' <summary>
''' Represents a text box control that performs configurable database queries
''' and displays matching records in a drop-down results window.
''' </summary>
''' <remarks>
''' The control supports configurable query columns, filtering, result selection,
''' frozen values, primary key tracking, and hyperlink-like interaction.
''' </remarks>
<Designer(GetType(QueriedTextBoxControlDesigner))>
<DefaultEvent("FrozenPrimaryKeyChanged")>
<DefaultProperty("Query")>
<DefaultBindingProperty("Query")>
Partial Public Class QueriedBox
    Inherits TextBox
    Friend DropDownResultsForm As DropDownResultsPopup
    Private WithEvents Timer As Timer
    Private _CtrlHyperlink As Boolean = False
    Private _IsHyperlink As Boolean = False
    Private _FirstEnter As Boolean = False
    Private _KeyDown As Boolean
    Private _Freezing As Boolean
    Private _PrimaryKeyAlias As String
    Private _RawFrozenValues As New List(Of (String, Object)) From {("Column", New Object())}
    <Category("QueriedBox")>
    Public Event FrozenPrimaryKeyChanged(sender As Object, e As FrozenPrimaryKeyEventArgs)
    <Category("QueriedBox")>
    Public Event FrozenPrimaryKeyChanging(sender As Object, e As FrozenPrimaryKeyEventArgs)
    <Category("QueriedBox")>
    Public Event HyperlinkClicked(sender As Object, e As FrozenPrimaryKeyEventArgs)
End Class
