Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Classe responsável por redimensionar controles ou formulários de maneira suave.
''' </summary>
Public Class FluidResizer
    Private _TargetSize As Size
    Private ReadOnly _Control As Control
    Private ReadOnly _ResizeTimer As Timer
    Private _IsResizing As Boolean = False ' Flag para evitar conflito de redimensionamento

    Public Sub New(Control As Control)
        _Control = Control
        _ResizeTimer = New Timer With {.Interval = 1}
        AddHandler _ResizeTimer.Tick, AddressOf ResizeTimer_Tick
    End Sub

    ''' <summary>
    ''' Define o tamanho final desejado para o controle e inicia o redimensionamento suave.
    ''' </summary>
    ''' <param name="TargetSize">O tamanho alvo que o controle deve alcançar.</param>
    ''' <remarks>
    ''' O redimensionamento é feito gradualmente em passos, com incrementos calculados 
    ''' para garantir uma transição suave. 
    ''' </remarks>
    Public Sub SetSize(TargetSize As Size)
        ' Verifica se o redimensionamento já está em andamento antes de iniciar outro
        If _IsResizing Then Return

        _TargetSize = TargetSize
        _IsResizing = True
        If Not _ResizeTimer.Enabled Then
            _ResizeTimer.Start()
        End If
    End Sub

    Private Sub ResizeTimer_Tick(sender As Object, e As EventArgs)
        Dim stepWidth As Integer = Math.Max(1, Math.Abs(_TargetSize.Width - _Control.Width) / 5)
        Dim stepHeight As Integer = Math.Max(1, Math.Abs(_TargetSize.Height - _Control.Height) / 5)

        If _Control.Width < _TargetSize.Width Then
            _Control.Width = Math.Min(_Control.Width + stepWidth, _TargetSize.Width)
        ElseIf _Control.Width > _TargetSize.Width Then
            _Control.Width = Math.Max(_Control.Width - stepWidth, _TargetSize.Width)
        End If

        If _Control.Height < _TargetSize.Height Then
            _Control.Height = Math.Min(_Control.Height + stepHeight, _TargetSize.Height)
        ElseIf _Control.Height > _TargetSize.Height Then
            _Control.Height = Math.Max(_Control.Height - stepHeight, _TargetSize.Height)
        End If

        If _Control.Width = _TargetSize.Width AndAlso _Control.Height = _TargetSize.Height Then
            _ResizeTimer.Stop()
            _IsResizing = False
        End If
    End Sub
End Class
