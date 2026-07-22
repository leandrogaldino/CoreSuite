Imports System.ComponentModel
Imports System.ComponentModel.Design

Partial Public Class QueriedBox
    Inherits TextBox
    Friend DropDownResultsForm As QueriedBoxDropDownResults
    Private _CtrlHyperLink As Boolean = False
    Private _IsHyperLink As Boolean = False
    Private _FirstEnter As Boolean = False
    Private _FreezedValue As String
    Private _UnFreezeColor As Color
    Private _IsFreezed As Boolean
    Private _FreezedPrimaryKey As Long
    Private _CharactersToQuery As Integer = 3
    Private _QueryInterval As Integer = 300
    Private _DropDownBorderColor As Color = SystemColors.HotTrack
    Private _GridBackColor As Color = SystemColors.Window
    Private _GridSelectionBackColor As Color = SystemColors.HotTrack
    Private _GridForeColor As Color = SystemColors.ControlText
    Private _GridSelectionForeColor As Color = SystemColors.Window
    Private _GridHeaderBackColor As Color = SystemColors.Window
    Private _GridHeaderForeColor As Color = SystemColors.ControlText
    Private _LabelBackColor As Color = SystemColors.Window
    Private _LabelForeColor As Color = SystemColors.ControlText
    Private _FreezeColor As Color = Color.Blue
    Private _DropDownAutoStretchRight As Boolean
    Private _DropDownStretchRight As Integer
    Private _QueryEnabled As Boolean = True
    Private _DesignerHost As IDesignerHost
    Private _KeyDown As Boolean
    Private _Freezing As Boolean
    Private _RawFreezedValues As New List(Of (String, String, Object)) From {("Table", "Field", New Object())}
    Private WithEvents Timer As Timer
    <Category("Propriedade Alterada")>
    Public Event FreezedPrimaryKeyChanged(sender As Object, e As EventArgs)
    <Category("Propriedade Alterada")>
    Public Event FreezedPrimaryKeyChanging(sender As Object, e As EventArgs)
    <Category("Ação")>
    Public Event HyperlinkClicked(sender As Object, e As EventArgs)
End Class
