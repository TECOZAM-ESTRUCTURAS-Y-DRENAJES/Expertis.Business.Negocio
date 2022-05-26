<Serializable()> _
Public Class DataPrcCrearPedidoCompraMantenimiento
    Public Preventivos() As DataOrigenPC
    Public IDContador As String
    Public IDOperario As String

    Public Sub New(ByVal Preventivos() As DataOrigenPC, ByVal IDContador As String, ByVal IDOperario As String)
        Me.Preventivos = Preventivos
        Me.IDContador = IDContador
        Me.IDOperario = IDOperario
    End Sub
End Class
