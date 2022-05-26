Imports System.Collections.Generic

Public Class PedCabEDI
    Inherits PedCab

    Public IDPedidoEDI As Integer
    Public IDPedido As Nullable(Of Integer)
    Public Mantener As Boolean
    Public DepartamentoEDI As String
    Public SeccionEDI As String
    Public IDDireccionEnvio As Integer

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
        Me.IDPedidoEDI = data("IDPedidoEDI")
        Me.Origen = enumOrigenPedido.EDI
        Me.Edi = True
        If Length(data("IDPedido")) > 0 Then
            Me.IDPedido = data("IDPedido")
        End If
        If IsDate(data("FechaPedido")) Then
            Me.Fecha = data("FechaPedido")
        End If
        If Length(data("IDCondicionPago")) > 0 Then
            Me.IDCondicionPago = data("IDCondicionPago")
        End If
        If Length(data("IDFormaPago")) > 0 Then
            Me.IDFormaPago = data("IDFormaPago")
        End If
        If Length(data("ObservacionesComerciales")) > 0 Then
            Me.ObsComerciales = data("ObservacionesComerciales")
        End If
        If Length(data("PedidoCliente")) > 0 Then
            Me.PedidoCliente = data("PedidoCliente")
        End If

        Me.Mantener = data("Mantener")
        If data.Table.Columns.Contains("Departamento") Then Me.DepartamentoEDI = data("Departamento") & String.Empty
        If data.Table.Columns.Contains("Sucursal") Then Me.SeccionEDI = data("Sucursal") & String.Empty

        If Length(data("IDDireccionEnvio")) > 0 Then Me.IDDireccionEnvio = data("IDDireccionEnvio")
    End Sub
End Class
