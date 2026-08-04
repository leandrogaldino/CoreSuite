Imports System.ComponentModel
Partial Public Class TextBoxActionPanel
    ''' <summary>
    ''' Occurs when the <see cref="TargetControl"/> reference changes.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when the target TextBoxBase reference changes.")>
    Public Event TargetControlChanged As EventHandler
    ''' <summary>
    ''' Occurs when the floating action panel becomes visible.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when the floating action panel becomes visible.")>
    Public Event PanelShown As EventHandler
    ''' <summary>
    ''' Occurs when the floating action panel is hidden.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when the floating action panel is hidden.")>
    Public Event PanelHidden As EventHandler
    ''' <summary>
    ''' Occurs when an enabled action is executed by mouse or through <see cref="PerformAction(String)"/>.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when an enabled text-box action is executed.")>
    Public Event ActionClicked As EventHandler(Of TextBoxActionClickEventArgs)
    Private Sub OnTargetControlChanged()
        RaiseEvent TargetControlChanged(Me, EventArgs.Empty)
    End Sub
    Private Sub OnPanelShown()
        RaiseEvent PanelShown(Me, EventArgs.Empty)
    End Sub
    Private Sub OnPanelHidden()
        RaiseEvent PanelHidden(Me, EventArgs.Empty)
    End Sub
    Private Sub ExecuteAction(Action As TextBoxAction)
        If Action Is Nothing OrElse Not _Enabled OrElse Not Action.Enabled OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed Then Return
        Dim E As New TextBoxActionClickEventArgs(Me, _TargetControl, Action)
        RaiseEvent ActionClicked(Me, E)
        Action.Invoke(E)
    End Sub
End Class
